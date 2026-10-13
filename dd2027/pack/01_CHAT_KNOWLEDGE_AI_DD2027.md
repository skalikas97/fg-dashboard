# 01 - CHAT KNOWLEDGE: AL-IDRISI drydock 2027, master dock book

Rewritten 13/10/2026 (second pass) from the claude.ai chat and Project "AI" (state file, case register, template analysis, memory). Written for a Claude Code session that starts with no memory of the chat. Read `00_START_HERE.md` first.

## 1. Goal and phases (read this before anything else)

- **Phase 1 (now): an EMPTY, PREDEFINED Excel template, ready for inputs, built faithfully on the TS's own reference workbooks.** The TS (13/10/2026): "on later phase i will put the real docking items for my docking but before i would like to have the excel template predefined to start inputs, and not change anything from the template i will use". Fidelity to his templates is the requirement. The build is a copy of their construction without their job content. Nothing is redesigned.
- **Phase 2 (later): the real docking items** are entered by the TS (and via the intake rules of his skill). Not part of the current build.
- **Deviations from his templates are OFF by default.** Anything that changes how a reference tab looks or works (status text column, Job ID formula, extra columns D1-D8, the fixes in file 03 A3, a Class tab, a Dashboard) is a proposal that needs his explicit yes. Only "input-ready" additions that do not change the look are recommended yes: empty pre-formatted buffer rows and empty Rev rows (file 03 A6 point 4).
- The TS gives his own instructions first, per tab. "Dont run" is a standing instruction: propose, wait for the answer, then build.

## 2. Who, what, when

- User: KoPsy (Konstantinos Psyroukis, initials PSKO), Technical Superintendent (TS), Jan De Nul Group (JDN) dredging fleet. Class society mainly Bureau Veritas (BV). Fleet Manager: Peter Broes (signs "Fleetcoordinator"; title to confirm).
- Vessel: AL-IDRISI ("AI"), trailing suction hopper dredger, 7,500 m3 hopper, built by STX Offshore & Shipbuilding (Korea), hull B-5055, IMO 9572707, call sign LXAI, BV 17705W, flag Luxembourg, owner CODRALUX SA, manager Dredging and Maritime Management S.A. Sister Vitus Bering (VB, B5057): its data is labelled context only.
- Event: AI 5-yearly drydock, target end March 2027, after the Barcelona dredging project. Handover: due 25/03/2027 (Thursday) [HO]; BV renewals due 26/03/2027 (Friday) [BV].
- Folder name (answered 05/10/2026): pattern is `YYYYMM - Place - Event`. Recommended `202703 - Spain - Drydock` or `202703 - Spain - Drydock (5Y renewal)`. "Term renewal" is not BV wording as far as known (unverified); a class renewal is five-yearly, not annual. "Spain" is a working title (the three yards that fit are Spanish). The old masterfile folder "202701 - Turkey - Drydock" is a predecessor name.

## 3. Method (fixed by the TS)

Per tab, three stages, TS sign-off between stages:
1. **OPTIMISE / SELECT** - compare how the reference files did it; options keep / merge / drop / new; recommend one with a short reason.
2. **CONSTRUCT** - empty frame only: title block, Rev block, legend, header, categories, widths, colours, validation, formulas, freeze, filter, print setup. In Phase 1 the frame IS the deliverable.
3. **CONTENT** - Phase 2: fill from sourced material with source tags; unsourced stays blank or "to be confirmed".

Also: one tab, one stage per turn unless he asks for more; never all tabs at once; ask only what the current tab needs, batched, each with a recommended default; do not restate earlier decisions; his corrections take effect at once and go into the state file the same turn. Construction conventions come only from the reference workbooks (file 03); their job content is ignored ("do not focus on the content of the works or the jobs, but on the way I construct the documents").

## 4. Working rules and TS preferences (apply always)

- **Confirm before writing:** before writing or integrating anything into his Excel master workbooks, show what will be written and get his explicit OK (stored preference; skill gate 9.0). A newly built template is a separate deliverable (file name with revision, Rev block row), not a write into a master.
- **Never load-and-save one of his reference originals** with openpyxl or LibreOffice for a deliverable (threaded comments, extension validation, rich text and cached values are lost). Read them, rebuild a NEW file from spec (file 03 A6).
- **Plain RGB, never theme references:** the reference themes differ per workbook (file 03 A4).
- **Document properties** of new files show the TS ("Psyroukis Konstantinos") as author and last editor or stay blank; never openpyxl or any tool name.
- Maritime, technical, regulatory, legal: flag uncertainty; no invented data or assumed specifics; source tags; reference numbers copied verbatim, never retyped from recollection.
- Symbols: hyphen (-) not em dash; plus-minus (+-) not tilde; arrow (->) not the arrow glyph. Dates dd/mm/yyyy in text, d-mmm in grids. No invisible characters in cells.
- Replies: short prose, what was built, what needs his decision; no recap of steps; tables only when they genuinely help. Drafts for class, yards, fleet managers, suppliers: formal and diplomatic. Emails and wording natural and humane, not stock phrases. Sign as PSKO.
- Think ahead of port calls and the dock: anticipate authorities, agent, suppliers, logistics, bunkering; propose items he may have forgotten for him to accept or reject.
- Evidence tags: `[HO]` handover, `[BV]` BV report, `[TS dd/mm]` stated by TS, `[W]` web; DERIVED = calculated (run in code); "Likely (not confirmed)" for hypotheses; conflicts shown with both values; TS statement governs over documents; newest message from the responsible party governs status.

## 5. Memory in Claude Code (replicate the Project as plain files)

Working folder (suggested): `reference/` (the 4 original .xlsx, read only), `pack/` (these .md files), `build/` (scripts), `out/` (deliverables with revision in the name).
- `STATE.md` = live status board (sections 6, 10, 11, 12 here are its content). Read first, update last, about 2 KB: tab board, decisions log (TS-confirmed only, dated, with the rejected option), open decisions, missing inputs, hard dates.
- `TEMPLATE_ANALYSIS.md` = file 03 (conventions, A1-A8) plus the Routing schema line per built tab.
- `CASE_REGISTER.md` = vessel facts with source and static/dynamic flag (file 02 holds the dock-relevant extract).
- The TS's saved skill `drydock-workbook-builder` encodes the same protocol (state first, load only what the stage needs, never re-extract the references, update state last, revision in file name and Rev block).

## 6. Status board (13/10/2026)

| Item | Status |
|---|---|
| Template tabs | None built. Stage 1 (optimise/select) is open for every tab. |
| SOW tab | Stage-1 proposal D1-D8 presented 29/09/2026 (section 7). Under the Phase 1 rule these are deviations: OFF until he says yes. Waiting for his own SOW instructions. |
| Yard list | Content draft only (frame not designed): `AI_DD2027_Yard_list_Rev02.xlsx` delivered 05/10/2026 (Rev00, Rev01 superseded). Phase 2 material; becomes the yard-comparison content later. |
| Pack for Claude Code | Files 00-03b delivered 13/10/2026 (second pass). |

### 6.1 Template decisions still to put to the TS (recommended default in brackets; he decides)

- T1 **Which reference tabs form the AI template, and from which file.** AA = execution (Daily SOW, Spares, Subcontractors-attendees, Planning); SO / SANTIAGO = commercial (SOW initial / final, Budget file, Costfile, AXI, Scaffolding, 3rd party Planner, 3rd parties). [All tabs of both roles in the reference order; he removes what he does not want.]
- T2 **One workbook or two.** [Two files mirroring his own practice: execution file (AA style) and commercial file (SO style). Skill leaves it to him.]
- T3 **Categories to keep as the pre-set rows.** AA 10 JDN groups for the execution tabs; SO letters A-X for the SOW tabs (file 03 A8). Full JDN code tree is not in the references (codes 300 and 600 missing): needed from him. [Keep each reference list as is.]
- T4 **Empty formatted rows per category** (does not change the look). [15 on Spares, 5 elsewhere, plus 10 empty Rev rows.]
- T5 **Placeholder text.** Title `AL-IDRISI - Drydock 2027 <Place>`; Rev00 PSKO; yard column titles "Yard 1" style. [As in file 03 A6 point 3.]
- T6 **Planning grid length.** Reference AA Planning has 35 day columns (27/10/2025 to 30/11/2025), 2.7 wide. First date and number of days for AI unknown. [Take the number of days and start date from him; dates driven by one start-date cell only if he accepts that as an input-ready addition.]
- T7 **Planner scroll bar** (SANTIAGO 3rd party Planner has a Display-week form control that openpyxl cannot write). [Plain "Display week" input cell with the same validation prompt, or leave the Planner out.]
- T8 **Style choices that only matter if he mixes AA and SO tabs:** Rev block order ("Rev. # | Name | Date" in AA, "Rev. # | Date | Name" in SO), font (AA Arial, SO Aptos Narrow), header grey, Ongoing colour (AA yellow, SO cyan), spacer column/row, tab colours, tab order. [When a tab copies one reference tab 1:1, its own style follows that tab and these questions do not arise.]
- T9 **Money tabs colour base.** "Not agreed" and "Obtained reduction" differ between SANTIAGO DOCKING and COSTFILE (file 03 A4). [Model on SANTIAGO DOCKING.]
- T10 Meaning of Dept letter "T" and the header "WL item" in the references; which cost-carrier code AI uses (references show 8578 / 8579, AI unknown).

## 7. SOW tab - the stage-1 proposal of 29/09/2026 (UNCONFIRMED; deviations, OFF by default)

Proposed columns: `Specs (JDN code) | Job ID (formula, unique) | Description | Dept | Afloat/Dock | Status (text, validated) | Class ref | Yard | 3rd party | Crew | JDN WS | Qty | Unit | WO item | Yard ref | RFQ | PO | Added in Rev | Remarks`. Transport file moved to Spares. Grouping by JDN system code, yard extras last.
- D1 tab names "Daily SOW" + "SOW initial" (his own names; "SOW final" = Daily SOW frozen at close). D2 Job ID formula over Specs (unique key, duplicate check). D3 status as validated text column. D4 status fill on Specs + Description by conditional formatting from the text. D5 Yard executor column. D6 Class ref column. D7 Added in Rev column. D8 transport file column to Spares.
- Reference construction to copy if he declines all of them: AA Daily SOW (file 03 A7, 03b) with status by manual fill on Specs + Description, no status column.
- Skill build-ins also proposed, also OFF until yes: Key dates input cells, Spares price/currency/total columns, cross-reference column for yard VO / EWO numbers, CLASS & SURVEYS, FOLLOW-UP and DASHBOARD tabs.

## 8. The TS's masterfile status convention (handover section 8)

For the existing AI drydock masterfile: yellow = ready for FM approval; orange = amount still to be determined; green = ready; no marking = still needs writing. The reference palette (file 03 A1) uses yellow for "ongoing/discussion" and orange for "TBC", so the two conventions overlap. Not a Phase 1 question; it matters when status fills are used in Phase 2.

## 9. Dates and voyage (context; newest governs)

- [TS 05/10] Barcelona project confirmed, project code 7755. AI released from Anaklia operations Fri 06/11 morning; last port call Poti; departure/clearance Sat 07/11 morning. Next bunker/repair day around Mon 26/10.
- [WK40, 28/09-04/10] Turkey pit stop for bunkering Mon 09/11 (also lub oil stock order; not in the 05/10 TS update, unconfirmed since); Barcelona ETA Sun 15/11 (preparation and maintenance week); project ETC Sun 22/11 with harbour master inspection 1-2 working days before; 5-yearly DD after Barcelona, due 26/03.
- [TS 30/09] AI ready to start dredging in Barcelona on 22/11/2026; all arrival logistics before that. DERIVED: 22/11/2026 is a Sunday; with a non-working weekend the inspection falls Thu 19/11 - Fri 20/11 at the latest, so arrival and logistics by about Wed 18/11 (authority calendar not verified). Departure 07/11 to 18/11 = 11 days; Poti -> Barcelona transit time not verified.
- Superseded (do not use): leave Anaklia 30/11 and start 1/12; mid-Nov start; "mid Oct / early Nov departure"; case register F-055/F-056 voyage wording.
- Dock: due 25/03/2027 [HO] / 26/03/2027 [BV]; DD moved from Dec 2026 to end March 2027, Barcelona first [TS]; the January in the old folder name is a leftover.
- Items due BEFORE the dock target (DERIVED from [BV]): SMC and SSC certificates expire Tue 09/03/2027 (issued 10/03/2022 and 31/03/2022 per the certificate table); ISM/ISPS renewal limit Tue 09/03/2027; Centre Tailshaft complete survey and CSL items 134-138 limit Wed 10/03/2027; ILO 152 periodical Thu 11/03/2027. That is 14-16 days before 25/03 (10/03 to 25/03 = 15 days). How they are covered is not stated in any source. "Centre Tailshaft" meaning for AI unconfirmed.
- Hull/Mach annual-for-renewal window Sat 26/12/2026 - Fri 26/03/2027; Bottom Survey in Dry Dock due 26/03/2027. Full table in file 02 B3.

## 10. Vessel facts relevant to the dock (Phase 2 content; full list in file 02)

- Particulars: LOA about 119.1 m; LBP 104.25 m; breadth 23.00 m; depth 10.75 m (BV) vs 6.30 m (GA, probably a different reference) - conflict; summer draught 8.15 m; GT 8,048 / NT 2,414; hopper 7,500 m3; DWT conflicting (12,555 / 11,800 / 11,637.294 t).
- Docking weights: light ship 5,169 t; empty displacement 5,919 t (= light ship + 750 t spares; mean draught 3.10 m, aft 5.04 m) [VIM Table 1-3 condition 1].
- Hull appendages for the block plan: two under-hull Schottel rudder propellers (SRP 3030 CP); hopper bottom doors protrude about 2.523 m below keel when open (4 double doors, 4.3 x 5.2 m free passage each); a docking plan is indispensable (dwg B0184000 not available).
- Machinery relevant to the dock: 2 x MAN 8L32/40 (72,000 h overhaul parts requested via AMOS), PS turbo overhaul, dredge pump gearbox bearing renewal, hydraulic pipeline leaks, 3 dredge valves above the hopper (external crane), Schottel PS seal; details in file 02 section C.
- **SB Schottel issue (dynamic, as of 02/10/2026):** lower gearbox SB rudder propeller; Lubmarine sample 825801 (28/08/2026) PQ index 20397 vs previous 26, Fe 71 vs 26, URGENT; second sample 10985504 (14/09/2026): PQ 382, Fe 39, still elevated. Week 40: extra samples requested (SEM-EDX, particle count, ferrography), reduced load, vibration/temperature OK. Source of particles unknown; likely candidate for opening in the dock, not confirmed.

## 11. Yard selection (done 05/10/2026, all UNCONFIRMED until written offers; Phase 2 content)

Screen rules: lift >= 5,919 t where a lift exists; dock length >= 119.1 m; clear width >= 23.0 m; appendage clearances, sill depth, block loading, displacement approval are yard approval items. TS conclusion (05/10): among six Spanish yards checked, only **Cernaval, Navantia Cadiz, Metalships Vigo** can take the vessel.

| Yard | Docks that fit (published) | Distance from Barcelona (nm, +-10%, own waypoints) | Return fuel cost at 11 kn (USD) |
|---|---|---|---|
| Cernaval, Campamento, San Roque | 2: graving 400 x 50 m; floating 230 x 39 m internal (lift conflict 35,000 t BPCO vs 10,000 t APBA) | 535 | 131,774 |
| Navantia Cadiz | 3: graving 386 x 66 m (official); graving 237 x 34.1 m and floating 246.2 x 41.3 m (directory only, lift unknown) | 600 | 147,784 |
| Metalships & Docks, Vigo | 1: floating 162 m overall (150 m over pontoon), 28.5 m clear, 10,000 t, max docked draught 6.7 m (official; replaces aggregator 24.8 m / 8,000 t) | 1,060 | 261,085 |
| MB92 Barcelona | 1: graving 220 x 36 m; yacht yard; dry-dock rehabilitation 2026-2028 - doubtful | 0 | 0 |
| Astilleros de Mallorca; Indunaval-Vatame | 0 (slipways 1,700 t; 3,500 t / 1,500 t) | 110; 135 | 27,094; 33,251 |

Also out: Navantia Cartagena (100 m), Cernaval Malaga (24 m beam). Named by Perplexity without dock data (to research): Chantier Naval de Marseille (about 180 nm), San Giorgio del Porto Genoa (355), Fincantieri Palermo (565), Palumbo Malta (715). Turkey not researched. No yard publishes a March 2027 slot calendar: availability NOT VERIFIED for all.

Fuel model (editable inputs in the xlsx): speed 11 kn (PLACEHOLDER; TS said lower than 14 kn); consumption 25 m3/day (TS upper limit); LSMGO density 0.85 t/m3 (assumption); LSMGO 1,530 USD/t (proxy: MABUX 05/10/2026 Gibraltar MGO LS; Barcelona price unknown); trip factor 2 (return, assumption). Time h = distance / speed; fuel m3 = h / 24 x consumption; t = m3 x density; cost = t x price x trip factor. Fuel only; no port dues, pilotage, towage. Screen formula in the xlsx column J: `=IF(AND(H<>"",H<$C$15),"NO - lift",IF(OR(F="",G=""),"width/length unknown",IF(AND(F>=$C$13,G>=$C$14),"OK on listed data","NO - size")))`.

Questions for the three shortlisted yards: which dock is offered; certified dock particulars (entrance and clear width, sill depth/max draught, permitted displacement, block capacity, keel-block height and adjustability); clearance for the 2.5 m hopper doors and the two Schottel units (needs the docking plan); dredger references; written March 2027 slot and outage plan; crane capacity and alongside depth; labour situation (Cadiz); ownership and which floating dock operates in March 2027 (Vigo); cost basis for a 2-4 week stay.

## 12. Decisions log (TS-confirmed only; newest last)

- 29/09/2026: method as in section 3; skill `drydock-workbook-builder` proposed and saved (v2, memory protocol).
- 30/09/2026: AI ready to dredge in Barcelona 22/11/2026.
- 05/10/2026: wants a possible-yards list (city, distance, fuel, approx. LSMGO mobilisation cost from Barcelona, docks that fit with dimensions, dock count and details). Sailing consumption about 25 m3/day maximum; speed lower than 14 kn. Verified yard data with Perplexity. Conclusion: only Cernaval, Navantia Cadiz, Metalships Vigo can take the vessel. Voyage update (section 9).
- 13/10/2026 (a): continues the build in Claude Code (credits when the chat 5-hour limit is exceeded); chat keeps the knowledge (Project "AI"), Code creates files; asked for this pack.
- 13/10/2026 (b): first build is an EMPTY PREDEFINED TEMPLATE ready for inputs; real docking items come in a later phase; "not change anything from the template i will use". Recorded as: template fidelity instead of redesign; deviations and weak-point fixes OFF unless he says yes. Pack re-done accordingly.

## 13. Missing inputs / open questions for the TS

1. Template decisions T1-T10 (section 6.1) and his own SOW instructions (blocking).
2. SOW decisions D1-D8 (only if he wants any deviation).
3. Full JDN code tree (codes 300 and 600 absent from the references).
4. Predecessor masterfile (folder "202701 - Turkey - Drydock\01 Specs and drawings\Scope of Work") and "AI - Preparations ER DD - Rev00.docx": not provided (Phase 2).
5. Docking plan dwg B0184000; planned passage speed; Barcelona LSMGO price; next project location after the dock; yard data for Marseille, Genoa, Palermo, Malta, Turkey; Perplexity source URLs (Phase 2).
6. Meaning of Dept "T", "WL item", "Centre Tailshaft" for AI, JDN codes 8038 / H8 / HPR-8023, AI cost carrier code.
7. Whether the Turkey bunker stop on 09/11 stands; how the March 2027 pre-dock class items (section 9) are covered.

## 14. Environment notes for Claude Code

- Originals are large (two are 5.7-5.9 MB, formatting to XFD / WVR). Never iterate `ws.max_column` / `max_row` blindly; cap scans (rows <= 800, columns <= 45). An earlier script was killed for memory.
- `freeze_panes` in openpyxl returns the stored scroll cell, not the split: read `ws.sheet_view.pane.ySplit` and `xSplit` (file 03 A7 has the verified values).
- Do not name a script `struct.py` (shadows the standard library). Run readers of untrusted files with `python3 -I`.
- Build with openpyxl; recalculate with LibreOffice headless on a scratch copy; check formula errors; render each tab to PDF/PNG and compare with the rendered original tab; keep files lean (format only the used range).
- Deliver each build with a revision in the file name and the Rev block (PSKO).
- `AI_DD2027_Yard_list_Rev02.xlsx` (Python/openpyxl, built in the chat) is the reference for the later Yard tab: blue-on-yellow input cells, screen formula, Rev block. The build script stayed in the chat container.

## 15. Chat chronology (TS messages, abbreviated)

1. Docking preparation request with the four reference workbooks. 2. "Do not focus on the content of the works or the jobs, but on the way I construct the documents." 3. Step-by-step plan and a memory-adaptive skill. 4. "Now we start with the tab SoW correct?" -> D1-D8. 5. 30/09 Barcelona 22/11/2026 line "for the memory". 6. Chat vs Code vs Cowork discussion; "I need first to give you instructions to make the SoW correctly. Dont run." 7. 05/10 voyage update; yard list requests; Perplexity round; "only those 3 can facilitate the vessel"; dock count and details. 8. 05/10 folder title; "class renewal could be also annual?". 9. 13/10 pack request. 10. 13/10 template-first instruction and "please check again thoroughly and optimize the files made for claude code".
