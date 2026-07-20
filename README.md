# DocWorks — GitHub Pages build (v0.9.1)

100%-local manufacturing document engine: drop a BOM (or drawing PDFs) and generate
Family Tree (F-), Parts List (P-), Travelers (TRV-), and Work Instructions (WI-).

## Deploy to GitHub Pages
1. Create a repository (e.g. `docworks`) and copy the contents of this zip to its root.
2. Commit and push.
3. Repo → Settings → Pages → Source: "Deploy from a branch" → Branch: `main`, folder `/ (root)` → Save.
4. Your app appears at `https://<user>.github.io/docworks/` within ~1 minute.

## Local testing
Opening `index.html` directly from disk (file://) mostly works now that the app is
precompiled, but for full fidelity serve it:
```
cd docworks && python -m http.server 8000
```
then open http://localhost:8000. GitHub Pages serves it correctly by default.

## Notes
- `index.html` loads React, SheetJS, and Babel from cdnjs at page load, and pdf.js /
  tesseract.js on first PDF/image drop. Your data never leaves the browser —
  network is used only to download library code. The future Tauri desktop build
  bundles all libraries for fully offline use.
- `samples/` contains three test BOM CSVs and three drawing PDFs — drag them in to try
  the full pipeline (drop all three DWG PDFs together to watch the BOM assemble).
- The app ships precompiled (app.js) — no in-browser compile step. app.jsx remains
  in the repo as the editable source; recompile with:
  `tsc --jsx react --allowJs --target es2020 --module none --outFile app.js app.jsx`
  then wrap the whole file in an IIFE -- `(() => { ... })();` -- so its top-level
  consts (e.g. XLSX) do not collide with library globals.
- Custom logic: Settings → Export templates JSON, edit, re-import (defaults untouched).
- Local LLM assist (optional, import-edge only): point Settings at an Ollama-style
  endpoint, default http://localhost:11434. Note: a page served over https cannot
  call plain-http localhost in some browsers — Chrome allows localhost as a secure
  origin; if blocked, run the desktop build or serve locally.

## Files
- index.html — loader shell
- app.jsx — the entire application (also the artifact source of truth)
- samples/ — BOM CSVs + drawing PDFs for testing
- DOCWORKS_Handoff_Brief.md — architecture/conventions reference
