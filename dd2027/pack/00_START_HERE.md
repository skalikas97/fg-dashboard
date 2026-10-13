# 00 - START HERE (Claude Code handover, AL-IDRISI drydock 2027)

Pack of 13/10/2026 (second pass). Owner: KoPsy (Konstantinos Psyroukis, PSKO), Technical Superintendent, Jan De Nul.

## What this is

Build the EMPTY, PREDEFINED Excel template for the AL-IDRISI 5-yearly drydock (target end March 2027), constructed exactly like the TS's own reference workbooks, ready for inputs. The TS's words: "not change anything from the template i will use". The real docking items come in a later phase. Job content in the references is ignored; only their construction is copied.

## Non-negotiable rules

1. **Fidelity first.** Reproduce the reference construction (columns, widths, fills, borders, merges, freeze, filter, print setup, folding, formulas). Any deviation (status text column, Job ID formula, extra columns, new tabs, fixes of weak points) is OFF until the TS says yes. Only empty formatted buffer rows and empty Rev rows are recommended additions, because they do not change the look.
2. **Confirm before writing.** Show what will be written, get his explicit OK. Never write into one of his masters. A new template is a separate file with a revision in the name and in the Rev block.
3. **Never load-and-save the 4 original workbooks** with openpyxl or LibreOffice for a deliverable. Read them; rebuild a NEW workbook from spec (file 03 A6).
4. **Write plain RGB**, never theme colours (themes differ per workbook, file 03 A4).
5. **Freeze panes:** read `sheet_view.pane.ySplit/xSplit`, not `ws.freeze_panes`.
6. **Document properties** of new files: the TS as author/last editor, or blank. No tool names.
7. **Accuracy:** flag uncertainty on maritime, technical, regulatory, legal points; no invented data; reference numbers copied verbatim.
8. **Symbols:** hyphen (-) not em dash; +- not tilde; -> not the arrow glyph. Dates dd/mm/yyyy in text, d-mmm in grids.
9. **One tab, one stage per turn**: optimise/select -> construct -> (later) content, with his sign-off between stages. "Dont run": propose first, wait for his answer.
10. **Replies:** short prose; what was built and what he must decide; recommended default for each open decision.

## Reading order

| File | Size | Read when |
|---|---|---|
| `01_CHAT_KNOWLEDGE_AI_DD2027.md` | about 21 KB | First, fully. Goal and phases, method, rules, status, template decisions T1-T10, dates, vessel and yard facts, decisions log. |
| `03_SAMPLING_REFERENCE_WORKBOOKS.md` | about 21 KB | Second, fully. Roles of the 4 references, conventions A1, colours A4 (RGB), what cannot be copied A5, clone method A6, per-tab figures A7, category lists A8. |
| `03b_TAB_DETAIL_GENERATED.md` | about 130 KB | Per tab only. Do not read in full. `grep -n "^### " 03b_TAB_DETAIL_GENERATED.md` lists the tabs; read the section of the tab being built (column table with widths, fills, borders, top block, categories, CF, DV, formulas). |
| `02_ATTACHMENTS_DIGEST.md` | about 20 KB | Phase 2 (content) or when a vessel, class or yard fact is needed. Not needed for the empty template. |

## First prompt to give Claude Code

"Read 00 and 01, then 03. Do not build anything. Confirm in 10 lines: the goal (empty predefined template, fidelity, no deviations unless I say yes), the rules, and the status. Then ask me only the template decisions T1-T10 from file 01 section 6.1, in one batch, each with its recommended default."

Next prompts, one tab at a time: "Tab <name>: stage 1 (optimise/select) from the references"; after his answer "stage 2 (construct the empty frame)"; deliver the file, wait for sign-off.

## Suggested folder

```
dd2027/
  pack/        00, 01, 02, 03, 03b (these files)
  reference/   the 4 original .xlsx, read only (do not save over them):
               AA_DAILY_SOW Afloat Repairs final.xlsx
               SO - SOW.xlsx
               SANTIAGO DOCKING.xlsx
               SO - COSTFILE - AXI SPLITUP.xlsx
  build/       scripts (do not name one struct.py; run readers with python3 -I)
  out/         deliverables, revision in the name (e.g. AI_DD2027_Execution_Rev00.xlsx)
  STATE.md     live status board (create from file 01 sections 6, 9, 12, 13; read first, update last, about 2 KB)
```

Also put `AI_DD2027_Yard_list_Rev02.xlsx` in `reference/` (content draft for the later yard tab, not a template).

## Technical reminders

- The two big originals (5.7-5.9 MB, formatting to column XFD / WVR): cap any scan at rows <= 800, columns <= 45.
- Build with openpyxl; recalc with LibreOffice headless on a scratch copy, zero formula errors; render original tab and new tab to PDF/PNG and compare; keep the file lean (format only the used range plus buffer rows).
- Not writable by openpyxl: threaded comments, form-control scroll bar, x14 conditional-format extension, hidden text box (file 03 A5). Ask the TS at that tab.
- Junk in the originals is not carried over: #REF! defined names, Vertex42 tags, external links, dead query table.
- Skill in the TS's account: `drydock-workbook-builder` (same protocol; its weak-point fixes are deviations here, OFF by default).
