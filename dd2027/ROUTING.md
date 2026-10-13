# Routing schema (one line per built tab; addendum to pack/03)

- X - Daily SOW - holds: jobs / scope items by yard letter category - not: spares, people, prices - key: Specs (col B), A.1 / A.1.1 - categories: A-X yard letters (15) - headers: Specs | Description | (Dept) | WL item | 3rd party | Crew | JDN WS | Remarks - Dept D/E/T - status: manual fill on Specs + Description, legend J2:J7 (Planned none / Ongoing FFFF00 / Completed 92D050 / TBC FFC000 / Cancelled/postponed FF0000 / All necessary parts on board 00B0F0) - dates: none in table; Rev dates and LastChange (H13, named) dd/mm/yyyy. Header row 15, data 16-105 (5 buffer rows per category), closing row 106.
- X - Spares / Subcontractors-attendees / Planning - same top block (Rev pulled from Daily SOW, LastChange); categories A-X; Spares 15 buffer rows, others 5; Planning day grid D17 start (Monday) -> 35 days, letters E/L/D/W/3 by CF.
- C - SOW initial (master, LastChange I13) / SOW final - AA top block, columns B-L as SO A-K; categories A-X, 5 rows each, header row 15.
- C - Costfile / Budget - per category 5 lines + "Subtotal <chapter>" row; total block; cost-code summary 0080/8038/xxx by SUMPRODUCT; names USD / EUR local (B4 / B5).
- C - AXI - company template; amounts linked to the Costfile summary rows; carrier 8038.
- C - Scaffolding (30 lines, A/D/N summary by SUMIF) / 3rd parties (6 vendor blocks) / 3rd party Planner (start date C4, Display week G4 + scroll bar, 8-week grid).
