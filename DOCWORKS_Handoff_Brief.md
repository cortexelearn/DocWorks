# DOCWORKS — Project Handoff Brief (v0.9.1)

(Formerly DOCWORKS — rebranded; internals and conventions unchanged.)

Purpose: allow any future Claude session (Opus or otherwise) to continue template
refinement without re-deriving design decisions. The artifact `docworks.jsx` is the
single source of truth — the PDF renderer evals its engine, so edits there flow to both.

## What this is
A 100%-local document engine for EZ Motors-style aerospace motor manufacturing.
Input: BOM (drag-drop Excel/CSV, or pasted indentured text from PDF).
Select build configuration → generates four controlled documents per job:
- **F-[PN]** Family Tree — drawing-style sheet set (title block, SVG tree, legend, rev history)
- **P-[PN]** Parts List — standalone indentured list document
- **TRV-[PN]** Travelers — one per Make assembly in bottom-up build order, with
  Qty Pass / Qty Rej per op, ATP record, disposition block
- **WI-[PN]** Work Instructions — mirrors traveler ops 1:1, adds sub-steps, amber
  DRAWING REFERENCE/SPEC blocks (blanks to complete at release), and placeholder
  SVG vignette renderings per op ("REPLACE WITH PROCESS PHOTO")

## Architecture (single file: docworks.jsx)
1. **Import layer** — `textToGrid`, `guessMapping` (fuzzy header detection),
   `gridToRows` (parent-column / level-column / flat modes), multi-file queue with merge.
2. **Data layer** — `buildBOM`: parts/children indexes + validation (orphans,
   duplicates, rev conflicts, cycles). Missing-BOM detection: leaf parts whose desc
   matches /assembl|assy|kit/ with no children → flag, resolvable as "purchased".
3. **Config layer** — `makeConfigs`: Full Actuator (with option checkboxes),
   Housed Motor, Frameless Set (two tops), Stator Only, Rotor Only. Tops resolved by
   description matchers, manual picker fallback.
4. **Ops library** — `OPS_LIBRARY`: routing templates matched by assembly description.
   Template ids: rotor, statorWinding (user's 13-step routing — authoritative),
   statorStack, motor (consumes completed STA + ROT; includes post-install electrical
   verification hold at Op 40 — user has not confirmed this op; ask), brake, gearhead,
   actuator (conditional ops based on options present), generic fallback.
   Each ops fn receives `(p, kids, bom, excluded, ref, refPN)`:
   - `ref(regex, fallback)` → "Description (PN, qty UOM)" from actual BOM children
   - `refPN(regex, fallback)` → bare PN
   These resolve lower-level part numbers into op text at generation time.
5. **Layout v2** — `layoutTree2` + `planSheets`: width-aware. Try plain horizontal →
   leaf ladders (stack ≥4 leaf children into 1-3 columns) → paginate into sheet set
   (collapse level-1 assemblies to "SEE SHEET n", each gets own sheet). Budgets:
   MAXW=980px, MAXH main 440 / sub 560 → PN lettering ≥ ~7.7pt on letter landscape.
6. **Vignette library** — `vignetteSvg(title)`: ~24 schematic SVG operation renderings,
   keyword-matched (PICK array; order matters — specific patterns before generic).
7. **Renderers** — TreeDoc (F), PartsListDoc (P), TravelerDocs, WIDocs.
8. **v0.9 additions** —
   - Template JSON layer: `exportTemplatesJSON()` (token-preserving: {PN} {DESC}
     {ref:regex|fallback} {refPN:regex|fallback}), `setCustomTemplates(decls)`,
     `activeLibrary()` — imported JSON templates shadow built-ins by id; default
     generation logic unchanged; per-op `onlyIf` regex + `autoNumber` supported.
   - `runDocumentCheck(bom, excluded, tops, purchased)` — in-app pre-print check
     (op sequence, holds, kit/stock, missing BOMs, unresolved ref() fallbacks).
   - Drawing module: `extractFromDrawingText(text)` — deterministic title-block PN
     + parts-list row extraction (3 row-pattern variants, OCR-suspect PN flagging);
     ImportPanel drawings mode (image drop → OCR adapter, PDF → pdfText adapter,
     paste path always available) feeding the standard Import Review, flat mode
     with parent prefilled.
   - Adapter hooks: `window.docworksAdapters (legacy alias motoflowAdapters still honored) = { tesseract: {recognize}, pdfText:
     {extract} }`; `llmNormalize(rawText, {url, model})` posts Ollama-style
     /api/chat expecting CSV back — import-edge assist ONLY, never generation.
   - StructureEditor: interactive tree correction (drag-drop reparent with cycle
     prevention, inline edit, add child, cascade delete, purchased toggle) →
     rebuilds via buildBOM(rows).
   - Project JSON export/import: {rows, srcLabel, activeCfg, excluded, purchased,
     meta:{wo,sn,prog,date,rev,eco,change}, customTemplates}.
   - Order Data: Doc Rev / ECO No. / Change Description → doc headers + rev history.
   - Settings panel: adapter status badges, Ollama URL/model + Test connection,
     templates export/import/reset, project export/import.

## Companion files (validation & PDF rendering, in /home/claude/val during sessions)
- `BOM_MOT-3000_Housed_Motor.csv`, `BOM_FRM-3500_Frameless_Set.csv`,
  `BOM_ROT-3120_Rotor.csv` — sample BOMs (also in project outputs)
- `validate.js` — harness: parse issues, config resolution, expected-template check
  (EXPECT array — extend when adding templates), op sanity (ascending numbers,
  ≥1 hold, kit-first, stock-last), missing-BOM, layout overlap. RUN AFTER EVERY
  TEMPLATE EDIT. It caught the frameless-stator template gap.
- `render.js` — evals engine from docworks.jsx, renders F/P/TRV/WI HTML → PDF via
  puppeteer (chromium at ~/.cache/puppeteer; F docs landscape, rest portrait).
- `layout2.js`, `wi_vignettes.js` — shared modules (also embedded in artifact).

## Conventions the user has specified (do not regress)
- Stator winding assembly ALWAYS has its own traveler (13 steps, exact order:
  kit → wind coils → insert coils → verify electrical rotation → connect leads →
  lace → form → pre-varnish test★ → varnish/impregnate → mfg cleanup/visual →
  final test★ → final QA★ → stock for NHA). Motor consumes it complete.
- Two documents per job: F-[top PN] and P-[top PN], cross-referenced.
- Traveler: Qty Pass / Qty Rej columns per op; disposition Qty Accepted/Rejected.
- Ops must reference actual lower-level PNs from the BOM (via ref/refPN).
- Family tree readable at letter size: ladders + pagination, never tiny boxes.
- Missing sub-BOM → amber flag, resolvable as PURCHASED; never silently ignored.
- Multi-file BOM drop merges sub-BOMs one file at a time (review queue).

## Refinement workflow (next phase, with real sample documents)
1. Load the matching sample BOM, generate the same document type.
2. Diff real vs generated. Classify each difference:
   template edit (OPS_LIBRARY text/holds/ATP rows) · layout tweak (renderer
   components) · new parameter field · matcher adjustment (description regexes
   in configs / ref() calls / OPS_LIBRARY match fns / OPTION_CLASSES).
3. Edit ONE template or component at a time; re-run validate.js pattern; regenerate.
4. Description-regex tuning is the most likely need (e.g., "Wire, Magnet" vs
   "Magnet Wire") — one-line changes at ref()/match() call sites.

## Practical tips for the next session
- Work in small increments; ask the model to edit specific sections, not regenerate
  the whole file. The file is ~large; use targeted string replacement.
- Preserve React hooks discipline: components called in loops must not use hooks
  (SheetDrawing/TreeDrawing are plain functions for this reason).
- Frameless set has two tops → doc numbers join with "+"; if user assigns a
  set-level PN, add it as a parent row for a clean single F-/P- number.
- Known open question: motor Op 40 (post-install electrical verification hold) was
  added on engineering judgment — confirm or remove per user practice.
