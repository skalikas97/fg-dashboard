# 03 - SAMPLING INPUT: construction map of the 4 reference workbooks

Revised 13/10/2026 (second pass: freeze panes corrected from the sheet XML, theme colours resolved to RGB, package features audited, cloning method added). Part B (per-cell detail, generated) moved to `03b_TAB_DETAIL_GENERATED.md`; read it per tab, only when building that tab.

**Purpose.** The TS wants the master dock book to start as an EMPTY, PREDEFINED TEMPLATE built exactly like his own reference files ("not change anything from the template I will use"). Real docking items come in a later phase. This file says how his reference files are constructed and how to reproduce a frame of them faithfully. Job content in the references is irrelevant and must not be copied.

**Reading order:** A0 (what the files are) -> A1-A3 (conventions) -> A4 (colours, use RGB) -> A5-A6 (what cannot be copied, how to clone) -> A7 (per-tab figures) -> A8 (category lists) -> 03b for the tab being built.

### A0. The four workbooks and their roles

| Code | File (as uploaded) | Role | Tabs, left to right |
|---|---|---|---|
| AA | AA_DAILY_SOW__Afloat_Repairs_final.xlsx | EXECUTION file of an afloat repair period (Afonso de Albuquerque, 2025; name inside the title reads "AA - Repairs 2025 Buenos Aires"); one workbook-wide revision, secondary tabs pull the Rev block from the master tab by formula | Daily SOW, Spares, Subcontractors-attendees, Planning |
| SO-SOW | SO - SOW.xlsx | COMMERCIAL SOW of a docking (Santiago barge, 2025) | SOW final (tab orange), SOW initial (tab light blue) |
| SO-DOCK | SANTIAGO DOCKING.xlsx | COMMERCIAL docking file, newest phase tab leftmost | Costfile - Invoice discussion, AXI cost breakdown, Scaffolding, 3rd party Planner, 3rd parties, SOW initial, Budget file |
| SO-COST | SO - COSTFILE - AXI SPLITUP.xlsx | Post-dock extract | AXI cost breakdown, Costfile - Invoice discussion |

Notes: SOW initial in SO-SOW and in SO-DOCK are identical in construction; AXI cost breakdown is identical in SO-DOCK and SO-COST; the two Costfile tabs differ slightly (Rev, freeze, a few cells). The generator marks identical tabs and does not repeat them.

Function of each tab (for the optimise stage): Daily SOW = living job list with responsibility matrix and status fill; Spares = one row per ordered article with REQ/PO/collo/transport/status and colour on the whole row; Subcontractors-attendees = people and service providers with request/granted status; Planning = day grid per job with team colours; SOW final / SOW initial = priced-scope list (qty, unit, Afloat/Dock, RFQ, PO, transport file) frozen at two moments; Budget file = yard price comparison; Costfile = invoice discussion (draft invoice vs agreed / discussion / not agreed / obtained reduction); AXI = pro-rata split of general yard cost into final EUR by cost code; Scaffolding = claim vs measured; 3rd parties = one block per vendor with travel/hotel/transport/working-time lines; 3rd party Planner = Gantt with TODAY line.

### A1. Construction conventions (from the template analysis, 29/09/2026)

#### Page anatomy (main tabs)
- AA: col A spacer (w 1-2), row 1 spacer (h 10.5). SO: starts at A1.
- Title block top-left, merged (AA B2:E8, SO A1:E5), 26-28 pt, 2 lines: "<VESSEL> - <Event> <Year> <Place>" / "<tab function>" (Scope of works, Overview of spares, Overview of subcontractors, General planning, Estimated/Actual scope of works). Grey fill; master tab distinguished (AA Daily SOW: dark red fill, white text; SO: red bold text on grey).
- Revision block right of title: AA "Rev. # | Name | Date", SO "Rev. # | Date | Name"; medium outer / thin inner border; one row per revision, appended downward; initials PSKO; dates mm-dd-yy. AA secondary tabs pull the rev block from Daily SOW by formula. Rev labels: "Rev09" (AA), "Rev 11" (SO SOW, AXI), "Rev. 11" (Costfile).
- Legend block right of rev block: coloured cells + label, no borders.
- Blank row, then single header row: grey fill (AA th0-0.35, SO th0-0.25), 14 pt, medium borders all round, centred, wrap, height 18.75-38. Freeze below header (verified 13/10 from the sheet XML split, table A7), autofilter on header.

#### Table body
- Column order: Specs (code) | Description | classifiers | references | execution/logistics | Remarks (last, widest 50-90).
- Category rows: "<code> <GROUP NAME>" (e.g. "000 GENERAL", "A. PILOTAGE, MOORING AND UNMOORING"), grey th0-0.25, bold, medium top/bottom; merged across table on Spares/Subcontractors.
- Data rows: AA Arial 10, SO SOW Aptos Narrow 11; dashed horizontal borders, medium vertical column separators, thin between Specs and Description.
- Devices: "x" responsibility matrix (3rd party / Crew / JDN WS); "/" = no reference; letter codes (Dept D/E/T; Afloat/Dock A/D/A+D bold; Y = booked); dark grey th2-0.50 cell = not applicable for that row type; parent/child rows (job row + person rows; parent code + sub-lines, parent code merged vertically in SOW initial); multiple refs stacked in one cell with line breaks.
- Block separators: thick right border after Transport File and after Remarks (SO SOW); narrow white spacer column between invoice block and decision block (Costfile).

#### Status and colour system
- Status is set by fill on key cells, no status value column: AA Daily SOW -> Specs+Description; AA Spares -> whole row; SO SOW -> Specs+qty; Costfile -> Draft Invoice USD; Subcontractors -> text "Granted/Requested/Canceled" coloured by conditional formatting, person name green/red, "Y" green.
- Palette: green 92D050 done/agreed/on board/granted; red FF0000 cancelled/too late/not completed; orange FFC000 TBC/close call; yellow FFFF00 ongoing (AA)/delivered in country/discussion/requested; cyan 00FFFF ongoing (SO); blue 00B0F0 all parts on board (AA); no fill = planned/not yet started/on the way.
- Collapse logic: finished lines grouped (outline level 1) so they fold away; cancelled lines grouped and hidden; unused tariff lines hidden (Costfile 287 rows). Spares: closed lines level 1, open lines level 0.

#### Planning (AA)
- Specs | Description | one column per day (w 2.7) | Remarks. Date header rotated 90, d-mmm, 8 pt; weekdays th2, weekends th2-0.25; double border after Sunday.
- Team fills: Engine room C6E0B4, Electricians B4C6E7, Deck FFFF00, Workshop F8CBAD, 3rd party FF8181; crew-change legend alongside. Letter codes E/L/D/W/3 in day cells auto-coloured by CF with font = fill (letter invisible). Description cells turn light yellow FFFFCC when the row has any planned day (COUNTA rule).

#### Money tabs (Budget file / Costfile)
- Plain title "PRICE COMPARISON - DRYDOCKING <VESSEL>" 13.5 pt; currency block rows 3-5 ("1 USD =" + rate, custom number formats with currency suffix, rate-source note in dark fill/white font); defined name EUR.
- Two-row header: row 6 group titles (Yard name / Draft Invoice / Agreed / Discussion / Not agreed / Obtained reduction / Final / Remarks / Cost Code) - neutral groups black fill white bold, decision groups in their status colour (Agreed 92D050, Discussion FFFF66, Not agreed th5+0.40, Obtained reduction th4+0.60), columns tinted full height; row 7 units (Par. | Description | Units | Qty. | USD/unit | USD).
- Cost code last column, red font. "Subtotal <chapter>" rows: label black fill white text, values grey, SUBTOTAL(9). Bottom: TOTAL WITHOUT DISCOUNT / DISCOUNT (rate in own cell) / GRAND TOTAL (yellow, bold) / Already paid / Remaining. Cost-code summary table below, formula-driven (SUMPRODUCT on cost code column); "with discount" column uses hardcoded 0.95.
- Column outline groups: budget block and secondary decision columns collapsible; %total hidden. Threaded comments = negotiation log. Footer &F left, &P/&N right.

#### Other tabs
- AXI: accounting template, formula cells in coloured font, instruction box + detail-code list at bottom.
- 3rd parties: one block per vendor "n) VENDOR (SCOPE)" grey bold bar, sub-header (no. | description | qty | unit | price/unit | total price | AVB | PO | DOCUMENTS | Remarks | ETCommence | ETComplete), fixed lines Air Travel / Hotel / Transport / Working time, "Total estimated cost"; Remarks used as checklist (Completed, Report received, Timesheet signed).
- Scaffolding: working sheet, no title block; claim vs measured side by side; A/D/N letter; AGREED / DISCUSSION / NOT AGREED summary.
- 3rd party Planner: Vertex42 Gantt (TODAY line, CF overdue red / done green / % done orange / planned blue); end-date formula contains #REF!.

#### Print, view, fonts
- A3 (paper 8), fit to width, landscape (SO, Planning, Subcontractors) or portrait (Daily SOW); narrow margins; print titles on header row (SO SOW); gridlines off (SO); zoom 80-90%.
- Fonts: AA Arial; SO SOW Aptos Narrow; Budget/Costfile Arial; AXI/Planner/3rd parties Calibri.

### A2. Differences between AA and SO styles (to settle for AI)
Rev column order and label format; fonts; header grey shade; "Ongoing" colour (yellow vs cyan); where status fill sits; spacer column/row; tab colours (SO only); tab order.

### A3. Observed weak points (reference files)
1. Status carried only as fill colour -> cannot be counted by formula.
2. Two key systems (function codes vs yard tariff letters).
3. Non-unique keys in AA Daily SOW (same code on two rows).
4. Discount 0.95 hardcoded in the cost-code summary; AXI tab linked to an external file.
5. Cost carrier differs between Budget (8578) and Costfile/AXI (8579).
6. Formatting stretched to column XFD/WVR -> Costfile sheet XML approx. 35 MB.
7. Merged cells inside data ranges limit sort/filter.
8. Group names differ slightly between tabs of the same workbook (AA: "SHIP CONSTRUCTION" vs "SHIPS CONSTRUCTION"; "HYDRAULIC SYSTEM" vs "HYDR SYSTEM") -> one master name list.
9. Junk carried inside the files: defined names that are #REF! (AED, CORRECT, FLANGE, totaalA, USD, _Toc...), Vertex42 tags, external links (1 in SO-SOW, 1 in SO-DOCK, 2 in SO-COST), a dead query table + deleted connection on the AXI tab, a hidden text box on the Planner. None of it is to be reproduced.
10. Theme dependence: colours stored as theme references resolve to DIFFERENT RGB per workbook (A4). Write RGB, never theme references, in the new file.

### A4. Theme colours resolved to RGB (verified from each workbook's theme part)

The reference files store many fills as theme references (thN+tint; index 0 = white, 1 = black, 2 = light grey, 3 = dark blue-grey, 4-9 = accents 1-6). The theme differs per workbook, so the same reference gives a different colour. Use the RGB values below, per workbook, and write plain RGB into the new file.

| Workbook | Theme (name; fonts) | Theme colour reference -> RGB |
|---|---|---|
| AA | Office; Calibri Light, Calibri | th0=FFFFFF; th0-0.05=F2F2F2; th0-0.25=BFBFBF; th0-0.35=A6A6A6; th2=E7E6E6; th2-0.25=AFABAB; th2-0.50=767171; th3+0.75=CCD4DE; th5+0.60=F8CBAD; th8+0.60=B4C7E7; th9+0.60=C5E0B4 |
| SO-SOW | Office; Aptos Display, Aptos Narrow | th0=FFFFFF; th0-0.25=BFBFBF; th0-0.35=A6A6A6; th9+0.60=B4E5A2 |
| SO-DOCK | Office 2007 - 2010; Cambria, Calibri | th0=FFFFFF; th0-0.15=D9D9D9; th0-0.25=BFBFBF; th0-0.35=A6A6A6; th1=000000; th3+0.40=558ED5; th3+0.75=B8D0ED; th4=4F81BD; th4+0.40=95B3D7; th4+0.60=B9CDE5; th5=C0504D; th5+0.40=D99694; th6+0.60=D7E4BD; th7+0.60=CCC1DA; th8+0.60=B7DEE8; th9+0.60=FCD5B5 |
| SO-COST | Office; Aptos Display, Aptos Narrow | th0=FFFFFF; th0-0.15=D9D9D9; th0-0.25=BFBFBF; th1=000000; th3+0.40=2D7FCF; th4=156082; th4+0.60=83CBEB; th5+0.40=F2AA84; th7+0.60=96DCF8 |

Consequence for the status colours of the money tabs: "Not agreed" (th5+0.40) is D99694 in SANTIAGO DOCKING (Office 2007-2010 theme) but F2AA84 in the COSTFILE extract (Office 2023 theme); "Obtained reduction" (th4+0.60) is B9CDE5 vs 83CBEB. Model the money tabs on SANTIAGO DOCKING unless the TS says otherwise. The exact (non-theme) palette is unaffected: 92D050, FF0000, FFC000, FFFF00, 00FFFF, 00B0F0, FFFF66, C6E0B4, B4C6E7, F8CBAD, FF8181, FFFFCC. Fonts: AA Arial (theme Calibri); SO-SOW Aptos Narrow (theme Aptos); SANTIAGO DOCKING tabs Arial / Calibri (its copy of SOW initial is Calibri, not Aptos Narrow).

### A5. Package features of the originals (what is NOT reproducible by load-and-save with openpyxl)

| Workbook | Features found in the package (13/10 audit) | Consequence |
|---|---|---|
| AA | 4 sheets, 4 printer-settings parts, 1 legacy comment (Subcontractors-attendees), 4 defined names, no external links, no threaded comments. Created by another user, last saved by PSKO. | Simple. Comment content is job content: leave out. |
| SO-SOW | 36 defined names (mostly #REF! and Vertex42), 1 external link, no comments, no conditional formatting. | Do not carry names or the link. |
| SO-DOCK (SANTIAGO DOCKING) | Costfile: 116 threaded comments (negotiation log) + legacy comments; AXI: legacy comments + a dead query table (connection deleted); 3rd party Planner: horizontal scroll-bar form control linked to G4 ("Display Week", 1-100), hidden text box, sheet protection (rows insert/delete and formatting allowed), x14 extension on conditional formatting, 12 CF ranges, 1 validation prompt on G4; 51 defined names (mostly #REF!), 1 external link; Costfile sheet XML is 35 MB because formatting runs to column WVR / row 1399. | Threaded comments, the scroll bar, the x14 CF extension and the text box cannot be written by openpyxl. Planner: ask the TS at that tab (options: plain "Display week" input cell with the same prompt, or leave the Planner out). Format only the used range. |
| SO-COST | 2 external links, threaded + legacy comments, same Costfile/AXI as above. | Same. |

Rule from the TS's own skill: an original that the TS has saved in Excel is his master. Never load-and-save it with openpyxl or LibreOffice for a deliverable (threaded comments, extension validation lists, rich text and cached values are lost). Read it, do not write it.

### A6. How to make the empty template (recommended method)

1. **Rebuild from spec, do not clone by load-and-save.** Create a NEW workbook with openpyxl. For each approved tab apply the spec from this file and 03b: column widths and groups, row heights, merges (title, legend, category rows only), fonts (explicit names), fills (RGB from A4), borders, alignment, number formats, freeze split (A7), autofilter range, zoom, gridlines, print setup (paper 8 = A3, orientation, fit to width, margins, print titles, footer `&F` / `&P/&N`), outline levels, conditional formatting (standard rule types), data validations.
2. **Clear the content, keep the frame.** Remove job lines, vendor names, yard names, dates and amounts. Keep: title block, Rev block, legend, header row(s), category rows, fixed company text (AXI instruction box and cost-place list, vendor-block labels, Subtotal/Total labels, summary formulas).
3. **Placeholders:** title `AL-IDRISI - Drydock 2027 <Place>` over the tab function line; Rev block with Rev00, PSKO and the build date; yard column title "Yard 1" style placeholder instead of the old yard name.
4. **Input-ready (does not change the look; proposals, recommended yes):** pre-formatted empty rows at the end of every category inside the filter, validation and CF ranges and the print area (skill default 15 on Spares, 5 elsewhere); Rev block with 10 pre-formatted rows; formulas that exist in the reference (SUBTOTAL, SUMPRODUCT summary, EUR/USD rate cells, AXI spread, Planner dates) rebuilt cleanly without hardcoded constants or external links.
5. **Verify each tab before delivery:** render the original tab and the new tab to PDF/PNG with LibreOffice and compare side by side (widths, fills, borders, wrapping, hidden/folded rows); recalc with LibreOffice on a scratch copy, zero formula errors; document properties show the TS (e.g. "Psyroukis Konstantinos") as author and last editor or stay blank, never openpyxl; no hidden sheets, no leftover comments, no defined names with #REF!.
6. **Fallback only if the render shows visible differences:** zip-level clearing of a COPY of the original (keep styles, theme, sheet XML structure; blank the cell values; rebuild sharedStrings; drop comments, threadedComments, persons, externalLinks, queryTables, connections, ctrlProps, drawings, calcChain; reset docProps; trim the Costfile XML to the used range). More faithful, much riskier; ask the TS before using it.
7. **Show-then-write:** before writing into any master workbook of the TS, show what will be written and wait for his explicit OK. A newly built template file is a separate deliverable (file name with revision, Rev block row), not a write into his masters.

### A7. Per-tab figures verified from the sheet XML and cells (13/10/2026)

"Frozen rows" is the true split (ySplit); the stored scroll cell in the files (e.g. A23, A670) is irrelevant.

| Tab | Header row | Frozen rows | Autofilter | Zoom | Gridlines | Print (orient / paper / fitW / fitH / scale) | Print titles | Hidden / outlined rows | Merges | CF rules | DV | Formulas |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| AA :: Daily SOW | 10 | 10 | A10:J111 | 85 | None | portrait / 8 / 0 / None / 80 | None | 16 / 22 | 1 | 0 | 0 | 0 |
| AA :: Spares | 10 | 10 | B10:S712 | 81 | None | portrait / 9 / None / None / None | None | 0 / 648 | 12 | 0 | 0 | 18 |
| AA :: Subcontractors-attendees | 10 | 10 | B10:P10 | None | None | landscape / 8 / None / 0 / 73 | None | 20 / 19 | 13 | 9 | 0 | 18 |
| AA :: Planning | 12 | 12 | None | None | None | landscape / 8 / None / 0 / 87 | None | 1 / 0 | 43 | 29 | 0 | 18 |
| SO-SOW :: SOW final | 7 | 7 | A7:K197 | 85 | False | landscape / 8 / None / 0 / 10 | $7:$7 | 0 / 0 | 5 | 0 | 0 | 3 |
| SO-SOW :: SOW initial | 7 | 7 | A7:K196 | 85 | False | landscape / 8 / None / 0 / 10 | $7:$7 | 0 / 0 | 25 | 0 | 0 | 0 |
| SO-DOCK :: Costfile - Invoice discussion | 6 | 6 | R8:R511 | 91 | False | landscape / 8 / None / 0 / 63 | None | 287 / 287 | 1 | 0 | 0 | 1037 |
| SO-DOCK :: AXI cost breakdown | 5 | 0 | A5:L5 | None | None | None / None / None / None / None | None | 0 / 0 | 2 | 0 | 0 | 173 |
| SO-DOCK :: Scaffolding | None | 0 | None | 105 | None | None / None / None / None / None | None | 0 / 0 | 2 | 0 | 0 | 57 |
| SO-DOCK :: 3rd party Planner | 7 | 7 | None | None | False | landscape / None / None / 0 / 47 | $4:$7 | 0 / 0 | 13 | 7 | 1 | 102 |
| SO-DOCK :: 3rd parties | 3 | 0 | None | 88 | None | portrait / 9 / None / None / None | None | 0 / 0 | 14 | 0 | 0 | 6 |
| SO-DOCK :: SOW initial | 7 | 7 | A7:K196 | 85 | False | landscape / 8 / None / 0 / 10 | $7:$7 | 0 / 0 | 25 | 0 | 0 | 0 |
| SO-DOCK :: Budget file | 6 | 6 | None | 85 | False | landscape / 8 / None / 0 / 71 | None | 3 / 0 | 1 | 0 | 0 | 582 |
| SO-COST :: AXI cost breakdown | 5 | 0 | A5:L5 | None | None | None / None / None / None / None | None | 0 / 0 | 2 | 0 | 0 | 173 |
| SO-COST :: Costfile - Invoice discussion | 6 | 6 | R8:R511 | 93 | False | landscape / 8 / None / 0 / 63 | None | 287 / 287 | 1 | 0 | 0 | 1037 |

Notes: Costfile and Budget freeze after the group-title row 6 (the units row 7 scrolls). AXI, Scaffolding and 3rd parties have no freeze. In SOW final/initial the print scale value 10 is overridden by fit-to-page. Costfile hides 287 unused tariff lines (outline level 1) and uses an autofilter on the Cost Code column only (R8:R511). Fold/hide counts show folding is a real part of the construction.

### A8. Category lists in the references (structure, not job content)

- AA (Daily SOW, Spares, Subcontractors-attendees, Planning): `000 GENERAL`, `070 SHIP CONSTRUCTION`, `100 MAIN DRIVES`, `200 AUX DRIVES`, `400 HYDRAULIC SYSTEM`, `500 ELECTRIC SYSTEM`, `700 WINCHES AND HOISTING GEAR SHIP CONSTRUCTION`, `800 AUXILIARY EQUIPMENT`, `900 DREDGING EQUIPMENT`, `1000 MISCELLANEOUS`. Codes 300 and 600 do not occur (the full JDN code tree is not in the references; ask the TS). Spares uses the same groups; Subcontractors-attendees too. Key format `NNN.N` / `NNN.NN.NN`.
- SO (SOW final / initial): `A. PILOTAGE, MOORING AND UNMOORING`, `B. DRYDOCKING`, `C. SERVICES`, `D. GENERAL`, `E. SURFACE PREPARATIONS`, `F. PAINT`, `G. STEEL`, `H. TANKS`, `I. SEAWATER LINES`, `J. HOPPER`, `M. PROPELLERS AND TAILSHAFTS`, `N. ENGINES`, `Q. AUXILIARIES`, `T. ANCHORS AND CHAINS`, `X. MISCELLANEOUS` (letters K, L, O, P, R, S, U-W are not used: the yard tariff skips them). Key format `A.1`, `B.2`.
- AXI cost breakdown (company accounting template, keep as is): cost places `1xx`, `00X`, `097`, `300`-`304`, `W & T Specific (only for exceptions)`, `400`-`409`, `M&R General`, `Misc Modifications`, `M&R Stone-dumping Equipment`, `M&R Remote Operated vehicles (ROV)`, `M&R Jack-up Equipment`, `M&R Trencher`, `430`, `M&R Cable layer`, `M&R Safety, Classification & Inspection`, `Security Vessels`, `Transport`, `Fuel`, `Fresh water supplies`, `Waste disposal & collection`, `Shore power supply`, then an instruction box. Columns: Costplace | Cost Carrier | Detailcode | Detail Description | Comments | Document / Curr | Initial Amount / Doc Cur | Discount Amount / Doc Cur | Final Amount / Doc Cur | Reallocation / General | Spread % | Final Amount / EUR.
- Costfile / Budget file: two-row header; chapter structure through `Subtotal <chapter>` rows and a GRAND TOTAL block; column groups listed in A1 and 03b.
- 3rd parties: vendor blocks `1)` ... `6)` each with sub-header (no. | description | qty | unit | price/unit | total price | AVB | PO | DOCUMENTS | Remarks | ETCommence | ETComplete).
- 3rd party Planner columns: DESCRIPTION | RESPONSIBLE | PLANNED START | PLANNED END | % DONE | WORK DAYS | DATE DONE | DAYS OVERDUE | day grid.
