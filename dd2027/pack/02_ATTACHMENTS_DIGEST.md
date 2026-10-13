# 02 - ATTACHMENTS DIGEST (what each source holds, and what the dock book can take from it)

Created 13/10/2026, revised 13/10/2026 (second pass: BV wording and dates fact-checked against the BV text). Source tags used below: [BV] BV MOVE survey status report generated 19/09/2026; [HO] TD handover DGJE -> PSKO 22/09/2026; [WK38] weekly report week 38; [WK40] weekly report week 40 (28/09-04/10/2026); [VIM] TD Manual AL-IDRISI / VITUS BERING Rev.2 (10-2013); [FM] JDN fleet TSHD manual rev.15 (23/08/2016); [PPX] Perplexity yard screening report, research cut-off 05/10/2026; [R:F-###] fact number in the vessel case register (dynamic facts must be re-verified before use). "NOT VERIFIED" is copied from the source and means no official source was found, not that the capability is absent.

## A. Inventory

| # | Attachment | Where it lives | Relevance to the dock book |
|---|---|---|---|
| 1 | AA_DAILY_SOW__Afloat_Repairs_final.xlsx (4 tabs) | chat upload | TEMPLATE (execution style) -> file 03 (+03b) |
| 2 | SO - SOW.xlsx (SOW final, SOW initial) | chat upload | TEMPLATE (commercial SOW) -> file 03 (+03b) |
| 3 | SANTIAGO DOCKING.xlsx (7 tabs, 5.9 MB) | chat upload | TEMPLATE (docking commercial file: costfile, AXI, scaffolding, planner, 3rd parties, budget) -> file 03 |
| 4 | SO - COSTFILE - AXI SPLITUP.xlsx (post-dock extract, 5.8 MB) | chat upload | TEMPLATE (AXI split + costfile) -> file 03 |
| 5 | 17705W-shipstatus 4.pdf = [BV], 21 pages | Project file | Class and statutory due dates, continuous survey list -> CLASS & SURVEYS tab |
| 6 | 22sept206 - TD Handover AI DGJE - PSKO.pdf = [HO], 10 pages | Project file | Drydock status, masterfile paths and colour convention, open technical items, logistics |
| 7 | AI - TD202640.pdf = [WK40] (weekly report week 40) | Project file | Latest voyage dates, drydock preparation status, subcontractors contacted |
| 8 | AI Weekly report wk38 = [WK38] | chat upload | Earlier status (superseded by HO / WK40 on conflict) |
| 9 | AI_Vessel_Certificates.pdf (1 page, created 24/09/2026) | Project file | Hull/Mach class renewal dates, O2 resuscitator inspection |
| 10 | Al-Idrisi.pdf (1 page spec sheet) | Project file | Rounded marketing datasheet (conflicts with BV on some figures) |
| 11 | GA STX.B5055-B0171000_-_0.pdf (image-only PDF) | Project file | General arrangement dwg B0171000 Rev.0 dated 25/11/2011, model 7500MSE01; only title block + principal particulars read |
| 12 | TD38 Suction hopper dredgers - deck.pdf (13 p), - engine room.pdf (63 p) | Project files | Fleet-wide spare part recommendations (rev 6.0, modified 15/01/2020; currency not confirmed) -> SPARES tab logic |
| 13 | List drawingsarticles AI - AL-IDRISI 8038 H8 HPR-8023.xlsx | Project file (blob, open with Python) | PLM drawing / article register (master data, NOT stock): "JDN.AI-" numbering, article numbers |
| 14 | AI_VI_manual.pdf = [VIM] (about 468 KB text), Manual Hopperdredgers rev.15 = [FM] (about 377 KB) | Project docs | Particulars, docking weights, equipment descriptions |
| 15 | Perplexity reply (242abd0d-attachment.txt) = [PPX] | chat upload | Yard docks screening -> section C |
| 16 | State file copy (91531e61-attachment.txt) | chat upload | Same content as the status board in file 01 sections 6, 9, 12, 13 |
| 17 | AI_DD2027_Yard_list_Rev00/01/02.xlsx + PDF | built in chat (Rev02 latest) | Yard comparison draft -> section D |
| 18 | "AI - Preparations ER DD - Rev00.docx", predecessor masterfile, dwg B0184000 | NOT PROVIDED | Missing inputs |

## B. Vessel and class data relevant to the dock

### B1. Particulars used for docking (register facts)
- LOA about 119.1 m [R:F-014]; LBP 104.25 m [F-015]; breadth 23.00 m [F-016]; depth 10.75 m (BV) vs 6.30 m "depth to main deck" (GA) CONFLICT [F-017]; summer draught 8.15 m [F-018]; GT 8,048 / NT 2,414 [F-021]; hopper 7,500 m3 [F-020].
- Deadweight: three conflicting values 12,555 t / 11,800 t / 11,637.294 t at 8.15 m (VIM, most technical) [F-024].
- Light ship 5,169 t; empty displacement 5,919 t (+750 t spares), mean draught 3.10 m, aft 5.04 m [VIM Table 1-3 condition 1] (the TS-supplied basis used in the yard screening).
- Accommodation 25 persons (BV) vs 21 on the 2011 GA drawing [F-025].
- Main engines 2 x MAN Diesel 8L32/40, 750 rpm; total 8,000 kW (BV) / 8,800 kW (spec sheet) [F-033]. Generators: 2 x INDAR 4,815 kVA, 1 aux Stamford 1,000 kVA, 1 emergency Stamford 150 kVA [F-036]. Shore supply 400 V/50 Hz/300 kVA [F-037].
- Under-hull: two Schottel rudder propellers, SRP 3030 CP (SB Lubmarine unit LM01714744; lower gearbox part list 1161419, drawing 1162044 idx 004, 14,300 kg; hull/side match of the part list NOT confirmed) [F-081, F-094].
- Hopper bottom doors: 4 double doors, 1 row; each section free passage 4.3 x 5.2 m; protrude 2.523 m below keel when open; shallow-water door 2.0 x 2.6 m; self-emptying line dia. 900 mm [F-042].
- Dredge pump IHC HR 222-47-110 WFT (impeller dia. 2.22 m), motor INDAR 4,000 kW; jet pump set Nijhuis 2 pumps/1 shaft; bow thruster 750 kW; hydraulics 300 bar [F-038 to F-041].
- Bow connection designed for floating line ID900 or ID750; installed female = ID750 per HO/WK38 (dynamic); ID900 spare in Ecuador [F-064, F-061].
- Draghead 2nd generation hydraulic, width 5.0 m; spare draghead onboard; ripping equipment kept onboard [F-044, F-048].
- Article register numbers for bow connection items exist (e.g. ID750 female complete 00546442, ID900 female complete 00546820) [F-067] - register is master data, not stock.

### B2. Class: Hull, machinery and notation [BV, F-026]
Class I Hull Mach, Hopper dredger, Unrestricted + Dredging > 15 mi from shore; additional class AUT-UMS, DYNAPOS-AM/AT. No Conditions of Class / Statutory Recommendations outstanding; no ISM/ISPS non-conformities; no MLC deficiencies (as of 19/09/2026) [F-030]. Exemptions in force: Load Line submersion; SOLAS Safety Equipment harmonised; COLREG Annex I-3(a) [F-031].

### B3. Class and statutory survey table (key rows from [BV] generated 19/09/2026, "Last / Due / Range / Status"; annual, intermediate, periodical and approval lines with no renewal date are left out - open the BV PDF for the full list)

| Survey | Last | Due | Range | Note |
|---|---|---|---|---|
| Hull - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2025 - 26/03/2027 | Within window |
| Hull - Annual for Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Hull - Annual | 07/02/2026 | | | |
| Hull - Intermediate | 01/02/2025 | | | |
| Underwater Bottom Survey | 23/10/2024 | | | |
| Bottom Survey in Dry Dock | 10/03/2022 | 26/03/2027 | 26/03/2027 | |
| Machinery - Renewal (CS) | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Machinery - Annual for Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Centre Tailshaft - Complete Survey | 10/03/2022 | 10/03/2027 | 10/03/2027 | Limit date in less than 6 months |
| PS Tailshaft - Complete | 23/10/2024 | 23/10/2029 | | |
| SB Tailshaft - Complete | 01/08/2023 | 01/08/2028 | | |
| Boilers - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Aux Oil-fired Boiler 1 / Exhaust Gas Boilers 1, 2 - Internal | 23/08/2024 | 23/08/2029 | | |
| Automation - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2025 - 26/03/2027 | Within window |
| DYNAPOS Renewal Survey | 11/03/2022 | 26/03/2027 | 26/12/2025 - 26/03/2027 | Within window |
| Load Line - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Saf. Construction - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2025 - 26/03/2027 | Within window |
| Saf. Equipment - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Launching App. - Quinquennial Test | 11/03/2022 | (no due shown) | | |
| Saf. Radio - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Oil Pollution / Sewage Pol / Air Pollution - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Ballast Water Mngt - Renewal | 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| HK Convention Renewal Survey; IHM Renewal Survey | 10/03/2022; 11/03/2022 | 26/03/2027 | 26/12/2026 - 26/03/2027 | |
| Anti-Fouling System - Renewal | 11/03/2022 | (no due shown) | | |
| ILO 152 Cargo Gear - Periodical | 11/03/2022 | 11/03/2027 | 11/03/2027 | Limit date in less than 6 months |
| ILO 152 Cargo Gear - Annual | 07/02/2026 | 07/02/2027 | 07/02/2027 | |
| ISM - Renewal; ISPS - Renewal | 10/03/2022 | 09/03/2027 | 09/12/2026 - 09/03/2027 | Limit date in less than 6 months |
| MLC Renewal inspection | 10/03/2022 | 19/04/2027 | 19/01/2027 - 19/04/2027 | |
| Service-supplier tests | AIS annual 06/02/2026; VDR annual 06/02/2026; LSA annual 05/02/2026; LSA renewal 10/03/2022 | | | |

Certificates relevant to the dock date [BV certificate table]: Safety Management Certificate (SMC) issued 10/03/2022, expires 09/03/2027; Ship Security Certificate (SSC) issued 31/03/2022, expires 09/03/2027; Class Certificate, Load Line, Safety Construction, Safety Equipment, Safety Radio, Ballast Water Management, Hong Kong Convention and EU IHM certificates all expire 26/03/2027; MLC Certificate (flag-issued, 15/04/2022) expires 19/04/2027. ISM Document of Compliance valid to 05/06/2028; ISM periodical audit (company) due 05/06/2027. The SMC and SSC are certificates that expire; the table rows "ISM - Renewal" and "ISPS - Renewal" (due 09/03/2027, range 09/12/2026 - 09/03/2027) are the audits that renew them.

Continuous survey items overdue or due in the next 6 months (all last by BV, due 10/03/2027): 26 Turning gear PS ME; 59 Turning gear SB ME; 69 Working air compressor; 70 Starting air vessels (3 ea); 71 Control air vessel; 78 Thermal oil expansion tank; 82 and 83 T.O. circulation pumps No.1/No.2; 119 and 120 L.O. coolers PS/SB ME; 125 HFO purifier; 126 MDO purifier; 128 L.O. purifier No.2; 134 Electric cables; 135 Megger test; 136 ER piping; 137 FO piping; 138 General maintenance survey at end of reclassification [BV, F-029, F-063]. The machinery Continuous Survey List is numbered up to item 138 (the last item); the BV text states no total.

Items due BEFORE the 25/03/2027 (Thursday) dock target (DERIVED from the dates above; weekdays computed): ILO 152 annual Sun 07/02/2027; SMC and SSC certificate expiry and ISM/ISPS renewal audits Tue 09/03/2027 (range from Wed 09/12/2026); Centre Tailshaft complete survey and the continuous survey items listed above, incl. 134-138, Wed 10/03/2027; ILO 152 periodical Thu 11/03/2027. That is 14-16 days before 25/03 for the March items. Coverage plan is not stated in any source -> ask the TS / BV.

Other class-related notes: [WK40] class renewal / 5-yearly DD (bottom survey in DD) 26/03/2027, "in window for hull, automation, DP and safety construction renewals"; hazardous materials inventory (Hong Kong convention): "next BV visit we need certificate for hazardous inventory"; remote Flag State Inspection awaiting date from Flag; flag derogation AIS until 08/10/2026 (new AIS on board, installation ongoing, derogation extension to be requested with Radio Holland commissioning confirmation); rescue boat davit winch brake near miss, repair planned in Barcelona with BV presence. O2 resuscitator inspection last 05/03/2026, due 25/09/2026 [CERT, F-032].

## C. Drydock-specific content of the handover and week 40 report

[HO] Section 8 Drydock (22/09/2026):
- AI needs to go into the 5-yearly drydock after the Barcelona project (final date end of March). Engine room has the biggest scope; orders for ER are placed (to be followed up). Deck needs to be checked.
- Preparation planning: "in the masterfile everything that is marked yellow is ready for FM approval. Orange the amount needs to be determined. Green is ready. No markings needs writing."
- Paths (company network, quoted as written): masterfile and ER scope `S:\Favorites\AI\06 Major repairs and docking\202701 - Turkey - Drydock\01 Specs and drawings\Scope of Work`; specifications `...\01 Specs and drawings\Specifications`; subcontractors contacted `...\03 Communication`.
- 8.1 Logistics: for ER a complete list exists ("AI - Preparations ER DD - Rev00.docx"); for deck most items are still in Ecuador, push there; check with the captain double.
- Technical items that feed the dock scope [HO 5.1, 5.2, 6]: SB Schottel metal particles; antifouling card (Mexico, stuck in customs; REQ 2038012165 released); ME overhaul parts through AMOS (pistons not renewed, 72k too early per manual); dredge pump gearbox bearings (all parts on board, Multigear offer 2026-30190 Rev.01); hydraulic pipeline leaks (technician to check the system in DD; "was performed on Cabral with good results"); ME oil delivery; earth fault power cable draghead (one phase damaged); 3 dredge valves above the hopper need overhaul with an external crane; turbo REQ 2038011961 overhauled, stuck at MAN ("also necessary for overhaul next DD"); water flaps REQ 2038012182 on order; local Hempel paint; spares for the dock mostly in Ecuador containers; 2 x 40 ft container in Aalst (waiting for green light to Barcelona); 4 AI spare containers + loose DD items in Ecuador (waiting approval) -> [WK40] transport approval acquired, 7 AI containers for DD to Belgium.
- [WK40] ER subcontractors contacted for: DP gearbox, ME generator, MAN 72k overhauls, MAN PS turbo overhaul, MAN Sacos (alarm system) upgrade, sewage treatment plant, Schottel PS seal replacement. Deck subcontractors contacted for: renewal of paint system (Hempel). "DD SoW ongoing". Other WK40 technical items: hydraulic pipeline leakages (P7 line renewed; new 900 ring and pin cylinders ordered for Barcelona), earth fault 400 VAC traced to submersible ballast pump, exhaust gas boilers temperature controller (new Gefran 850 ordered), turning gland SB renewed, overflow SB fore movable part thickness measured.
- TD schedule [WK40]: PSKO until 25/10, Jelle de Groote from 25/10.

The handover and WK40 text above is project content, not template. In the dock book it belongs in the content stage (SOW lines, SPARES, 3RD PARTIES), after the frames are approved.

## D. Perplexity yard screening [PPX] (research cut-off 05/10/2026, no source URLs supplied in the pasted text)

Basis given to Perplexity: twin-screw TSHD; LOA 119.1 m; beam 23.0 m; moulded depth 10.75 m; light ship 5,169 t; empty displacement about 5,919 t; two under-hull Schottel rudder propellers; hopper bottom doors protrude about 2.5 m below keel when open; planned docking March 2027, 2-4 weeks, departing Barcelona.

Screening rules used by the report: LOA and beam alone do not make a facility suitable; docking displacement, arrival draught, sill depth, stability, block loading, appendage clearances and access remain approval items; "width" is not assumed to be entrance width or clear internal width unless explicit; lift capacity matters for floating docks, syncrolifts, shiplifts and slipways; a yard-specific docking plan is indispensable for the hopper doors and Schottel units (no yard publishes block adjustability to 2.5 m or door-opening pockets); no yard publishes a March 2027 slot calendar.

| Yard | Facilities and published figures (as reported) | Not verified |
|---|---|---|
| MB92 Barcelona | Panamax graving dock 220 x 36 m; shiplift 4,800 t (yachts up to 115 m LOA); syncrolift 2,000 t (up to 70 m); yacht-dedicated (35-200 m); 2026-2028 investment over EUR 40 million incl. rehabilitation of the 220 m dry dock | Entrance width, clear width, sill depth, permitted displacement, block capacity, acceptance of a commercial dredger, outage dates |
| Astilleros de Mallorca, Palma | Four slipways, max 1,700 t; STP travelifts up to 1,000 t, berths up to 90 m; outfitting quays 120 m | Anything for haul-out of AI: not viable on lift |
| Indunaval-Vatame, Burriana | Slipway 1 up to 120 m LOA and 3,500 t light ship (directory: 120 x 25 m); slipway 2 up to 90 m and 1,500 t; custom docking beds | Not viable on lift; afloat work possible |
| Cernaval, Campamento, San Roque | Graving dock 400 x 50 m; floating dock 230 x 39 m internal (BPCO: 35,000 t; APBA: 10,000 t; port article: 45 m external beam, ships up to 36 m beam); lay-by berths 510 x 50 m and 240 x 40 m; contacts comercial@cernaval.com, +34 956 699 314 | Entrance/clear width of the graving dock, sill depth, allowable displacement, block heights, dredger reference, March 2027 slot, lift conflict |
| Navantia Cadiz | Three docks incl. one floating; largest graving dock 386 x 66 m, 440,000 dwt (official); directory only: 237.0 x 34.1 m graving and 246.2 x 41.3 m floating; more than 3,000 m quay at 9 m draught, cranes up to 100 t; contact reparaciones@navantia.es, +34 981 35 20 40 | Entrance width, sill depth, blocks, lift of the floating dock, slot; summer 2026 workload gap and 2025 strike; "Costa Atlantica" dredger precedent doubtful |
| Metalships & Docks, Vigo | Floating dock 162 m overall (150 m over pontoon), 35.2 m external, 28.5 m clear, max docked draught 6.7 m, 10,000 t, ship limit 27.5 m breadth and 25,000 dwt; covered slipways 1+2 joined 150 x 38 m 3,500 t (no); 700 m berth, quay cranes 25 tm; contact metalships@metalships.com, +34 986 811 800 | Keel-block adjustability, actual docking displacement, ownership transfer (not completed at cut-off), which floating dock operates in March 2027 (a 9,500 t floating dock was processed by Vigo Port in September 2025) |

Decision summary of the report: strongest candidates on published physical data are Cernaval Algeciras, Navantia Cadiz, Metalships Vigo, Chantier Naval de Marseille, San Giorgio del Porto / Ente Bacini Genoa, Palumbo Malta, Fincantieri Palermo; Metalships is the only yard whose official data clearly combine adequate nominal lift (10,000 t), length and clear width, with the 6.7 m draught still to confirm. Marseille, Genoa, Palermo and Malta came with no dock data. Contact details above are as reported by Perplexity from official pages; they were not independently verified.

## E. Yard list workbook (built in the chat)

`AI_DD2027_Yard_list_Rev02.xlsx`, sheet "Yard list" (Rev block Rev00, Rev01, Rev02, PSKO, 05/10/2026; 56 formulas, 0 errors on recalculation).
- Inputs C6-C10 (blue on yellow): speed 11 kn (placeholder), consumption 25 m3/day, LSMGO density 0.85, LSMGO price 1,530 USD/t (MABUX indication 05/10/2026 Gibraltar MGO LS proxy), trip factor 2. AI particulars C13-C15: LOA 119.1, beam 23.0, empty displacement 5,919 t.
- Header row 19. Columns: # | Yard | City/location | Docks that fit (no.) | Dock details | screen basis (dock length / clear width / lift / max docked draught) | Screen vs AI | Distance (nm) | Steaming time (h) | Fuel one way (m3) | Fuel (t) | LSMGO cost one way (USD) | total (x trip factor) | Data status | Risks / to ask the yard | Contact.
- Screen formula: see file 01 section 10. Distances are own waypoint estimates (haversine over waypoints) +-10%; fuel only.
- Results are in file 01 section 10.

## F. Other facts from the case register that matter in the dock (all dynamic unless stated)

- Fuel consumption: dredging about 29.5 m3/day (HFO -> MGO switch in Anaklia); alongside/anchorage about 3.8 m3/day MGO [HO, F-057]. TS says sailing consumption about 25 m3/day maximum.
- Bunkering: 450 MT MGO on 28/08/2026; about every 3 weeks to Poti by truck (expensive port call: combine with repairs) [WK38, HO].
- Open items at handover and since: SB Schottel (file 01 section 9); ME oil supply (Aurelia TI 4020 8,320 L + Visga 32 120 L; 44 barrels from Istanbul to Poti) outcome not stated; AIS derogation; PLC batteries REQ 2038011927; SART resolved; antifouling card REQ 2038012165 stuck in customs.
- Barcelona project mobilisation (separate matter, Project "AI" hub-and-spoke): male part of the bow coupling, ID800 hose interface, Ecuador parts, start date. Not part of the dock book, but the dock follows it.
- Register conflicts still open: deadweight, accommodation, depth, speed/power, JDN codes 8038/H8/HPR-8023, Barcelona start date, pump-ashore parts location, Fleet Manager title, ID800 vs 750/900 interface, SB Schottel severity wording, ME oil outcome [case register section 3].
