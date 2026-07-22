const { useState, useMemo, useCallback, useRef, useEffect } = React;
const XLSX = window.XLSX;

/* =====================================================================
   DOCWORKS — drawing-style family tree (readable letter-size sheet sets) · multi-file import · missing-BOM resolution
   [ Import layer ] drag-drop xlsx/xls/csv/tsv · paste CSV or indentured
                    text (from PDF copy) · column mapping + preview
   [ Data layer   ] canonical rows → parts/children indexes + validation
   [ Config layer ] build configurations scope the BOM by matchers
   [ Ops library  ] routing templates keyed to assembly type
   [ Renderers    ] familyTree / traveler / workInstruction
   ===================================================================== */

/* ---------------- Sample data ---------------- */
const SAMPLE = `Parent,Find,Qty,UOM,PartNumber,Rev,Description,MakeBuy,Material,Remarks
,,1,EA,ACT-1000,1,Electromechanical Actuator Assembly,Make,,
ACT-1000,10,1,EA,GH-2000,A,Planetary Gearhead Assembly,Make,,
ACT-1000,20,1,EA,MOT-3000,C,BLDC Motor Assembly,Make,,
ACT-1000,30,1,EA,BRK-4000,B,Brake Assembly,Make,,
ACT-1000,40,1,EA,ENC-5000,A,Encoder Assembly,Buy,,
ACT-1000,50,1,EA,HSG-6000,A,Main Housing,Make,,
ACT-1000,60,1,EA,OUT-7000,A,Output Shaft,Make,,
ACT-1000,70,1,EA,CON-8000,A,Harness,Make,,
GH-2000,10,1,EA,GH-2100,A,Gear Housing,Make,,
GH-2000,20,1,EA,SHA-2101,A,Sun Gear,Make,,
GH-2000,30,3,EA,PLN-2102,A,Planet Gear,Make,,
GH-2000,40,1,EA,CAR-2104,A,Planet Carrier,Make,,
GH-2000,50,1,EA,RNG-2105,A,Ring Gear,Make,,
GH-2000,60,2,EA,BRG-6203,A,Bearing,Buy,,
MOT-3000,10,1,EA,LAM-3110,B,Stator Lamination Stack Assembly,Make,,
MOT-3000,20,1,EA,WND-3112,B,Three-Phase Winding,Make,,
MOT-3000,30,1,EA,ROT-3120,A,Rotor Assembly,Make,Various,Balanced / Ground
MOT-3000,40,2,EA,BRG-6002,A,Rotor Bearing,Buy,440C Stainless Steel,Matched / Approved Source
LAM-3110,10,120,EA,LAM-3111,A,Electrical Steel Lamination,Buy,M19,Insulated
LAM-3110,20,1,EA,ADH-3112,A,Bonding Epoxy,Buy,High Temp,Bond Stack
ROT-3120,10,1,EA,SHA-3121,A,Rotor Shaft,Make,17-4PH Stainless Steel,Inspect journals and datums
ROT-3120,20,8,EA,MAG-3122,A,Permanent Magnet,Buy,"NdFeB, high-temperature grade",Polarity-controlled matched set
ROT-3120,30,1,EA,SLV-3123,A,Magnet Retention Sleeve,Make/Buy,Nonmagnetic high-strength alloy,Bonded over magnet OD
ROT-3120,40,AR,GM,ADH-3124,A,Magnet Bonding Epoxy,Buy,Structural high-temperature epoxy,"Record lot, mix, cure"
ROT-3120,50,AR,GM,ADH-3125,A,Sleeve Bonding Epoxy,Buy,Structural high-temperature epoxy,"Record lot, mix, cure"
BRK-4000,10,1,EA,COI-4101,A,Brake Coil,Make,,
BRK-4000,20,1,EA,BAC-4103,A,Brake Back Iron,Make,Low Carbon Steel,Cd Plate
BRK-4000,30,1,EA,ARM-4102,A,Armature Plate,Make,,
BRK-4000,40,3,EA,SPR-4104,A,Compression Spring,Buy,,
BRK-4000,50,1,EA,FRI-4105,A,Friction Disc,Buy,,`;

// Example of an indentured parts list as copied out of a formatted PDF
const SAMPLE_INDENTURED = `Level  Part Number  Rev  Description                          Qty  UOM
0      ACT-1000     1    Electromechanical Actuator Assembly  1    EA
1      GH-2000      A    Planetary Gearhead Assembly          1    EA
2      GH-2100      A    Gear Housing                         1    EA
2      SHA-2101     A    Sun Gear                             1    EA
2      PLN-2102     A    Planet Gear                          3    EA
2      CAR-2104     A    Planet Carrier                       1    EA
2      RNG-2105     A    Ring Gear                            1    EA
2      BRG-6203     A    Bearing                              2    EA
1      MOT-3000     C    BLDC Motor Assembly                  1    EA
2      LAM-3110     B    Stator Lamination Stack Assembly     1    EA
3      LAM-3111     A    Electrical Steel Lamination          120  EA
3      ADH-3112     A    Bonding Epoxy                        1    EA
2      WND-3112     B    Three-Phase Winding                  1    EA
2      ROT-3120     A    Rotor Assembly                       1    EA
3      SHA-3121     A    Rotor Shaft                          1    EA
3      MAG-3122     A    Permanent Magnet                     8    EA
3      SLV-3123     A    Magnet Retention Sleeve              1    EA
3      ADH-3124     A    Magnet Bonding Epoxy                 AR   GM
3      ADH-3125     A    Sleeve Bonding Epoxy                 AR   GM
2      BRG-6002     A    Rotor Bearing                        2    EA
1      BRK-4000     B    Brake Assembly                       1    EA
2      COI-4101     A    Brake Coil                           1    EA
2      BAC-4103     A    Brake Back Iron                      1    EA
2      ARM-4102     A    Armature Plate                       1    EA
2      SPR-4104     A    Compression Spring                   3    EA
2      FRI-4105     A    Friction Disc                        1    EA
1      ENC-5000     A    Encoder Assembly                     1    EA
1      HSG-6000     A    Main Housing                         1    EA
1      OUT-7000     A    Output Shaft                         1    EA
1      CON-8000     A    Harness                              1    EA`;

/* =====================================================================
   IMPORT LAYER
   ===================================================================== */

/* --- generic delimited-line splitter (csv/tsv, quote-aware) --- */
function splitDelim(line, delim) {
  const out = []; let cur = "", q = false;
  for (const ch of line) {
    if (ch === '"') { q = !q; continue; }
    if (ch === delim && !q) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur); return out.map(s => s.trim());
}

/* --- text → raw grid. Detects tsv, csv, or column-aligned text --- */
function textToGrid(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return { grid: [], format: "empty" };
  const tabs = lines.filter(l => l.includes("\t")).length;
  if (tabs > lines.length / 2) return { grid: lines.map(l => splitDelim(l, "\t")), format: "tsv" };
  const commas = lines.filter(l => l.includes(",")).length;
  if (commas > lines.length / 2) return { grid: lines.map(l => splitDelim(l, ",")), format: "csv" };
  // column-aligned / space-delimited (typical PDF copy-paste): split on runs of 2+ spaces,
  // falling back to single-space tokenization when that yields one column
  let grid = lines.map(l => l.trim().split(/\s{2,}/).map(s => s.trim()));
  const oneCol = grid.filter(r => r.length === 1).length;
  if (oneCol > lines.length / 2) grid = lines.map(l => l.trim().split(/\s+/));
  return { grid, format: "text" };
}

/* --- header auto-detection: fuzzy match column names to canonical fields --- */
const FIELD_DEFS = [
  { key: "level",  label: "Level",       re: /^(level|lvl|indent|ind\.?)$/i },
  { key: "parent", label: "Parent PN",   re: /^(parent|nha|next higher|assembly|parent part)/i },
  { key: "find",   label: "Find / Seq",  re: /^(find|seq|sequence|item\s*no\.?|line|pos)/i },
  { key: "qty",    label: "Qty",         re: /^(qty|quantity|qty\s*per|per\s*assy)/i },
  { key: "uom",    label: "UOM",         re: /^(uom|unit|um|u\/m)$/i },
  { key: "pn",     label: "Part Number", re: /^(part\s*(number|no\.?|#)?|component|item(\s*id)?|material\s*no|pn)$/i },
  { key: "rev",    label: "Rev",         re: /^rev/i },
  { key: "desc",   label: "Description", re: /^(desc|description|part\s*name|name|nomenclature)/i },
  { key: "mb",     label: "Make/Buy",    re: /^(make|buy|make\/buy|m\/b|source|purch)/i },
  { key: "mat",    label: "Material",    re: /^mat(erial|l)?\.?$/i },
  { key: "rem",    label: "Remarks",     re: /^(remark|note|comment)/i },
];
const PN_RE = /^[A-Za-z0-9][A-Za-z0-9\-_./]{2,}$/;
const UOM_SET = new Set(["EA", "GM", "G", "KG", "LB", "OZ", "FT", "IN", "M", "MM", "CM", "ML", "L", "PC", "PCS", "SET", "AR"]);

function guessMapping(grid) {
  const first = grid[0] || [];
  const map = {}; // fieldKey -> column index
  let hasHeader = false;
  first.forEach((cell, ci) => {
    for (const f of FIELD_DEFS) {
      if (f.re.test(cell) && map[f.key] === undefined) { map[f.key] = ci; hasHeader = true; break; }
    }
  });
  if (!hasHeader) {
    // No header row: infer by content of first data rows
    const sample = grid.slice(0, Math.min(8, grid.length));
    const nCols = Math.max(...sample.map(r => r.length));
    const colIs = (ci, pred) => sample.every(r => r[ci] === undefined || r[ci] === "" || pred(r[ci]));
    for (let ci = 0; ci < nCols; ci++) {
      if (map.level === undefined && colIs(ci, v => /^\d{1,2}$/.test(v) && +v <= 12)) { map.level = ci; continue; }
      if (map.pn === undefined && colIs(ci, v => PN_RE.test(v)) && sample.some(r => /[-]/.test(r[ci] || ""))) { map.pn = ci; continue; }
      if (map.rev === undefined && colIs(ci, v => /^[A-Z0-9]{1,3}$/i.test(v))) { map.rev = ci; continue; }
      if (map.qty === undefined && colIs(ci, v => /^(\d+|AR)$/i.test(v))) { map.qty = ci; continue; }
      if (map.uom === undefined && colIs(ci, v => UOM_SET.has(v.toUpperCase()))) { map.uom = ci; continue; }
      if (map.desc === undefined && colIs(ci, v => /\s/.test(v) || v.length > 8)) { map.desc = ci; continue; }
    }
  }
  return { map, hasHeader };
}

/* --- level column → parent inference --- */
function levelsToParents(rows) {
  const stack = []; // stack[level] = pn
  const issues = [];
  return {
    rows: rows.map((r, i) => {
      let lvl = parseInt(String(r.level).replace(/\D/g, ""), 10);
      if (isNaN(lvl)) { issues.push(`Row ${i + 1}: unreadable level "${r.level}" — treated as level 1.`); lvl = 1; }
      const parent = lvl === 0 ? null : stack[lvl - 1] || null;
      if (lvl > 0 && !parent) issues.push(`Row ${i + 1} (${r.pn}): level ${lvl} has no level ${lvl - 1} above it.`);
      stack[lvl] = r.pn; stack.length = lvl + 1;
      return { ...r, parent };
    }),
    issues,
  };
}

/* --- commit a mapped grid to canonical rows --- */
function gridToRows(grid, map, hasHeader, hierarchyMode, flatParent) {
  const body = hasHeader ? grid.slice(1) : grid;
  const get = (r, k) => (map[k] !== undefined && r[map[k]] !== undefined) ? String(r[map[k]]).trim() : "";
  let rows = body
    .filter(r => get(r, "pn"))
    .map((r, i) => ({
      parent: hierarchyMode === "parent" ? (get(r, "parent") || null) : hierarchyMode === "flat" ? (flatParent || null) : null,
      level: get(r, "level"),
      find: get(r, "find") || String((i + 1) * 10),
      qty: get(r, "qty") || "1",
      uom: get(r, "uom") || "EA",
      pn: get(r, "pn"),
      rev: get(r, "rev") || "-",
      desc: get(r, "desc") || "",
      mb: get(r, "mb") || "",
      mat: get(r, "mat") || "",
      rem: get(r, "rem") || "",
    }));
  let convIssues = [];
  if (hierarchyMode === "level") {
    const conv = levelsToParents(rows);
    rows = conv.rows; convIssues = conv.issues;
  }
  if (hierarchyMode === "flat" && flatParent) {
    // ensure the parent itself exists as a row so the tree roots correctly
    if (!rows.some(r => r.pn === flatParent)) {
      rows.unshift({ parent: null, find: "", qty: "1", uom: "EA", pn: flatParent, rev: "-", desc: flatParent + " (imported single-level BOM parent)", mb: "Make", mat: "", rem: "" });
      rows = rows.map(r => r.pn === flatParent ? r : { ...r, parent: flatParent });
    }
  }
  return { rows, convIssues };
}

/* =====================================================================
   DATA LAYER — canonical model + validation
   ===================================================================== */
function buildBOM(rows, extraIssues = []) {
  const issues = [...extraIssues];
  const parts = {}, children = {};
  for (const r of rows) {
    if (!r.pn) { issues.push("Row with missing part number skipped."); continue; }
    if (!parts[r.pn]) parts[r.pn] = { pn: r.pn, rev: r.rev, desc: r.desc, mb: r.mb, mat: r.mat, rem: r.rem };
    else {
      if (parts[r.pn].rev !== r.rev && r.rev !== "-") issues.push(`Revision conflict on ${r.pn}: "${parts[r.pn].rev}" vs "${r.rev}".`);
      if (!parts[r.pn].desc && r.desc) parts[r.pn].desc = r.desc;
    }
    if (r.parent) (children[r.parent] = children[r.parent] || []).push(r);
  }
  for (const p of Object.keys(children)) if (!parts[p]) issues.push(`Parent "${p}" is referenced but has no row of its own.`);
  for (const [p, kids] of Object.entries(children)) {
    const seen = {};
    for (const k of kids) { if (seen[k.pn]) issues.push(`Duplicate child ${k.pn} under ${p}.`); seen[k.pn] = 1; }
    kids.sort((a, b) => (parseInt(a.find) || 0) - (parseInt(b.find) || 0));
  }
  const isChild = {}; for (const kids of Object.values(children)) for (const k of kids) isChild[k.pn] = 1;
  const tops = Object.keys(parts).filter(pn => !isChild[pn]);
  function hasCycle(pn, stack) {
    if (stack[pn]) { issues.push(`Circular reference at ${pn}.`); return true; }
    stack[pn] = 1; for (const k of (children[pn] || [])) if (hasCycle(k.pn, stack)) return true; delete stack[pn]; return false;
  }
  tops.forEach(t => hasCycle(t, {}));
  return { rows, parts, children, tops, issues };
}
function parseCsvBOM(text) {
  const { grid } = textToGrid(text);
  const { map, hasHeader } = guessMapping(grid);
  const { rows, convIssues } = gridToRows(grid, map, hasHeader, map.parent !== undefined ? "parent" : "level", null);
  return buildBOM(rows, convIssues);
}

/* =====================================================================
   CONFIGS + OPS LIBRARY (unchanged engine)
   ===================================================================== */
const OPTION_CLASSES = [
  { key: "gearhead", re: /gearhead/i, label: "Gearhead" },
  { key: "brake", re: /brake assembly/i, label: "Brake" },
  { key: "sensor", re: /encoder|sensor|resolver/i, label: "Encoder / Sensor" },
  { key: "cable", re: /harness|cable/i, label: "Cable / Harness" },
];
function makeConfigs(bom) {
  const findByDesc = re => bom ? Object.values(bom.parts).find(p => re.test(p.desc)) : null;
  return [
    { id: "actuator", name: "Full Actuator Assembly", desc: "Complete unit: motor, gearhead, brake, sensor, housing, harness. Options selectable.", hasOptions: true,
      tops: () => bom && bom.tops.length ? [bom.tops[0]] : null },
    { id: "housed", name: "Housed Motor Assembly", desc: "Motor assembly built and tested as a deliverable unit.",
      tops: () => { const m = findByDesc(/motor.{0,28}(assembly|assy)|(assembly|assy).{0,28}motor/i); return m ? [m.pn] : null; } },
    { id: "frameless", name: "Frameless Motor Set (Stator + Rotor)", desc: "Matched stator and rotor shipped as a component set — no housing or bearings.",
      tops: () => { const s = findByDesc(/stator|lamination stack/i), r = findByDesc(/rotor assembly/i); return s && r ? [s.pn, r.pn] : null; } },
    { id: "stator", name: "Stator Only", desc: "Lamination stack and winding through electrical test.",
      tops: () => { const s = findByDesc(/stator|lamination stack/i); return s ? [s.pn] : null; } },
    { id: "rotor", name: "Rotor Only", desc: "Shaft, magnets, sleeve — bonded, ground, balanced.",
      tops: () => { const r = findByDesc(/rotor\s+ass(?:embl|y)/i); return r ? [r.pn] : null; } },
  ];
}

const SAFETY_TEXT = {
  magnets: [["WARNING", "NdFeB magnets are brittle and develop strong attractive forces. Handle one at a time; keep magnets separated; keep ferrous tools clear. Keep magnets and magnetized rotors away from pacemakers, implants, magnetic media, and precision instruments."]],
  epoxy: [["CAUTION", "Epoxy adhesives: gloves and eye protection per SDS; observe pot life; ventilate per SDS."]],
  grind: [["CAUTION", "Magnet grinding dust is flammable and contaminating. Never dry-grind NdFeB without approved extraction."]],
  electrical: [["WARNING", "Dielectric and functional testing use hazardous voltages / energized rotation. Only trained, authorized personnel; use released test setups and guards."]],
};

const OPS_LIBRARY = [
  { id: "rotor", match: p => /rotor\s+ass(?:embl|y)/i.test(p.desc),
    atp: [["Magnet OD / sleeve OD size", "Per drawing"], ["Runout / concentricity", "Per drawing"], ["Magnet polarity pattern", "Alternating, per map"], ["Surface finish", "Per drawing"]],
    safety: ["magnets", "epoxy", "grind"],
    ops: (p, kids, bom, excluded, ref, refPN) => [
      { op: "10", dept: "STORES / QA", title: "Kitting", hold: true, text: `Kit ${ref(/shaft/i, "rotor shaft")}, matched ${ref(/permanent magnet|magnet$/i, "magnet set")}, ${ref(/sleeve/i, "retention sleeve")}, ${ref(/magnet bonding/i, "magnet bonding adhesive")}, ${ref(/sleeve bonding/i, "sleeve bonding adhesive")}, masking, and consumables per ${p.pn} BOM and kitting form. Verify shelf life and traceability.`,
        sub: [`Pull all items against the released ${p.pn} BOM; verify part number, revision, and quantity.`, `Verify ${refPN(/permanent magnet|magnet$/i)} matched-set identification and polarity labels.`, `Verify ${refPN(/magnet bonding/i)} and ${refPN(/sleeve bonding/i)} lot numbers and expiration dates — reject expired material.`, "Complete the kitting form and attach material certifications."],
        accept: "BOM and kitting form compliant; material/lot numbers recorded.", record: "Kitting form, lots, certs, QA stamp." },
      { op: "20", dept: "QA / INSPECT", title: "Receiving / Pre-Assembly Inspection", hold: true, text: "Inspect shaft journals, magnet seats, datums, sleeve ID/OD, magnet dimensions, coating condition, and polarity identification before assembly.",
        sub: ["Dimensionally verify shaft journals and datum surfaces per drawing.", "Inspect each magnet for chips, cracks, and coating damage under magnification.", "Verify sleeve ID/OD and condition."],
        accept: "Conforms to released drawings; no chips, cracks, corrosion, or coating damage.", record: "Inspection results, inspector stamp/date." },
      { op: "30", dept: "ASSEMBLY", title: "Bond Surface Preparation", text: `Clean and prepare ${ref(/shaft/i, "shaft")} and ${refPN(/permanent magnet|magnet$/i)} magnet bond surfaces per approved process. Mask journals and critical datums.`,
        sub: [`Install ${refPN(/shaft/i)} shaft in soft-jaw fixture; mask both journals and datum faces — verify full coverage.`, "Clean magnet seat area with approved solvent and lint-free wipes; single-direction wipes, one use per face.", `Clean bond face of each ${refPN(/permanent magnet|magnet$/i)} magnet; keep magnets segregated, one per tray position, labels visible.`, "Verify surfaces dry and residue-free under bright light; no bare-hand contact after cleaning.", "Record solvent lot and completion time — bonding must start within the released surface-active window."],
        photo: "Shaft in preparation fixture, journals masked, magnet seat cleaned.",
        accept: "Surfaces clean, dry, damage-free; masking covers all journals and datums.", record: "Solvent lot, method, completion time, operator." },
      { op: "40", dept: "ASSY / QA", title: "Magnet Dry-Fit & Polarity Verification", hold: true, text: `Dry-fit ${ref(/permanent magnet|magnet$/i, "magnets")} in the polarity fixture; verify sequence, pole orientation, axial location, and spacing.`,
        sub: ["Install polarity/assembly fixture; verify fixture ID and condition.", "Dry-fit all magnets in build sequence; verify full seating with no rock.", "Gauss-check alternating N-S pole pattern; complete the polarity map on the traveler.", "Verify axial location against fixture reference per drawing.", "Apply orientation witness reference.", "Call QA — do not mix adhesive until dry-fit and polarity map are accepted."],
        photo: "All magnets dry-fitted; gauss meter verifying alternating polarity; witness mark visible.",
        callout: { k: "CAUTION", v: "Never place two loose magnets within attraction range of each other or the assembled rotor — uncontrolled attraction fractures magnet corners." },
        accept: "Polarity map correct; full seating; axial location per drawing. QA accepts before adhesive mix.", record: "Polarity map, fixture ID, QA stamp/date." },
      { op: "50", dept: "ASSEMBLY", title: "Magnet Bonding", text: `Mix and apply ${ref(/magnet bonding/i, "magnet-bonding epoxy")} per released process; bond ${refPN(/permanent magnet|magnet$/i)} magnets in controlled sequence; remove excess without disturbing bondline.`,
        sub: [`Verify ${refPN(/magnet bonding/i)} lot/expiration; mix per released ratio; record mix start time and pot life.`, "Bond one magnet at a time to its mapped position; apply released adhesive pattern — no starved or flooded joints.", "Seat fully and confirm orientation unchanged from dry-fit.", "Complete all magnets within pot life; if exceeded, stop and mix fresh (record).", "Remove excess adhesive before gel; keep pole faces, journals, and datums adhesive-free."],
        photo: "Adhesive pattern on magnet bond face; fillet condition after seating.",
        accept: "All magnets seated in mapped positions; bondlines continuous; no adhesive on pole faces or datums.", record: "Adhesive P/N, lot, expiration, mix ratio/time, window, operator." },
      { op: "60", dept: "ASSY / QA", title: "Magnet Adhesive Cure & Bond Inspection", hold: true, text: "Fixture and cure per released schedule; monitor time/temperature; inspect all bonds after cure.",
        sub: ["Install cure fixture; verify magnets remain seated under load.", "Cure per released schedule; record start/stop, oven ID; retain chart where required.", "Post-cure: inspect each bondline under magnification — continuous fillet, no voids beyond allowance, no movement, no cracks/chips.", "Re-verify polarity map with gauss meter.", "Call QA for bond acceptance before grinding."],
        photo: "Acceptable cured fillet vs. rejectable void/starved joint — workmanship comparison.",
        accept: "Cure within released schedule; bondlines conform; polarity re-verified.", record: "Cure start/stop, oven ID, chart ref, results, QA stamp." },
      { op: "70", dept: "MACHINE", title: "Grind Magnet OD", text: "Grind magnet OD to drawing size and runout between released datums. Approved coolant and dust collection; protect journals; do not overheat.",
        sub: ["Set up between released datums; verify setup runout before grinding.", "Confirm extraction and coolant operating before wheel contact.", "Rough grind leaving finish stock; finish to drawing size, runout, and finish.", "Light passes only — stop on any discoloration, smearing, or chatter.", "Clean thoroughly; remove all magnetically-held grinding fines."],
        photo: "Rotor between centers on grinder; coolant and extraction in place; journals protected.",
        accept: "Magnet OD size, runout, taper, finish per drawing; no thermal damage.", record: "Final OD, runout, taper, finish, equipment ID." },
      { op: "80", dept: "QA / INSPECT", title: "Post-Grind Inspection", hold: true, text: "Inspect ground OD and rotor geometry; verify no magnet cracking, chipping, bond separation, overheating, or datum damage.",
        accept: "Dimensions and visual condition accepted before sleeve installation.", record: "Inspection results, inspector stamp/date." },
      { op: "90", dept: "ASSEMBLY", title: "Retention Sleeve Installation", text: `Clean ground OD and ${ref(/sleeve$|retention sleeve/i, "sleeve")} ID; apply ${ref(/sleeve bonding/i, "sleeve-bonding epoxy")} per released process; install sleeve with approved fixture and controlled press/thermal method; maintain axial location.`,
        sub: [`Clean magnet OD and ${refPN(/sleeve$|retention sleeve/i)} sleeve ID; verify dry and residue-free.`, `Verify ${refPN(/sleeve bonding/i)} lot/expiration; mix and record.`, "Apply adhesive per released pattern — avoid starvation and trapped air.", `Install ${refPN(/sleeve$|retention sleeve/i)} with approved fixture; controlled press or thermal method per released process.`, "Verify axial position and uniform squeeze-out at accessible edge; remove excess before gel."],
        photo: "Sleeve installation with alignment fixture; uniform adhesive squeeze-out at sleeve edge.",
        accept: "Sleeve fully seated at drawing position; uniform bond evidence; no starvation.", record: "Adhesive lot/mix, method, sleeve position, time, operator." },
      { op: "100", dept: "ASSY / QA", title: "Sleeve Adhesive Cure", hold: true, text: "Fixture and cure per released schedule; verify sleeve seated and concentric during cure; inspect bond edges after cure.",
        accept: "Cure within schedule; sleeve position and bond condition conform.", record: "Cure data, oven ID, results, QA stamp." },
      { op: "110", dept: "MACHINE", title: "Finish Grind Sleeve OD", text: "Finish grind sleeve OD to drawing on rotor datums. Control temperature and force — thin-wall sleeve over bonded magnets. Protect journals; clean after grinding.",
        accept: "Sleeve OD, runout, taper, finish per drawing; no damage.", record: "Sleeve OD, runout, taper, finish, equipment ID." },
      { op: "120", dept: "QA / INSPECT", title: "Final Inspection", hold: true, text: "Final dimensional and visual inspection: journal sizes, overall length, magnet/sleeve axial position, sleeve OD, runout, concentricity, finish, workmanship. Verify traveler complete.",
        photo: "Final inspection — rotor in V-blocks, indicator on sleeve OD.",
        accept: "All drawing characteristics accepted; inspection report attached.", record: "Inspection report; final QA stamp/date." },
      { op: "130", dept: "QUALITY / STOCK", title: "Preserve, Identify & Stock", text: "Clean, preserve, identify, and package. Protect journals and sleeve OD; package for magnetic handling. Release to next higher assembly or ship.",
        accept: "Preservation, ID, and packaging per released requirements.", record: "Qty accepted, stock location, lot/serial, final QA approval." },
    ] },
  { id: "statorWinding", match: p => /stator/i.test(p.desc) && /assembl|assy/i.test(p.desc) && !/lamination|stack|core/i.test(p.desc),
    atp: [["Phase resistance U-V / V-W / W-U", "Per ATP at stated temp"], ["Resistance imbalance", "Per ATP"], ["Insulation resistance", "Per ATP"], ["Surge comparison", "Per ATP"], ["Dielectric withstand", "No breakdown"], ["Line-to-line inductance", "Per ATP"], ["Electrical rotation / phase sequence", "Correct per drawing"]],
    safety: ["epoxy", "electrical"],
    ops: (p, kids, bom, excluded, ref, refPN) => [
      { op: "10", dept: "STORES / QA", title: "Kit Parts for Stator", hold: true,
        text: `Kit all components for ${p.desc} (${p.pn}) per BOM, including the completed and accepted ${ref(/lamination stack/i, "Stator Lamination Stack Assembly")}, ${ref(/magnet wire/i, "magnet wire")}, ${ref(/insulation/i, "slot insulation")}, ${ref(/leadwire|lead wire/i, "leadwire")}, ${ref(/lacing/i, "lacing materials")}, and ${ref(/varnish|resin/i, "impregnation varnish")}. Verify wire lot/gauge, stack acceptance status, and shelf-life items.`,
        sub: [`Pull all items against the released ${p.pn} BOM; verify part number, revision, and quantity.`, `Verify ${refPN(/lamination stack/i)} traveler is complete and QA-accepted.`, `Verify ${refPN(/magnet wire/i)} gauge, insulation class, and lot certification.`, `Verify ${refPN(/varnish|resin/i)} and consumable shelf life — reject expired material.`, "Complete the kitting form and attach material certifications."],
        accept: "BOM and kitting form compliant; stack accepted; lots recorded.", record: "Kitting form, wire lot, stack traveler ref, certs, QA stamp." },
      { op: "20", dept: "WIND", title: "Wind Magnet Wire Coils",
        text: `Wind coil groups from ${ref(/magnet wire/i, "magnet wire")} per released winding specification: turns count, wire gauge, strands in hand, and coil geometry per drawing.`,
        sub: ["Verify winding specification revision and coil data before starting.", "Set up winder with released tooling; verify tension per spec.", "Wind coil groups to released turns and geometry; count verification per released method.", "Protect finished coils from damage and contamination; identify coil groups."],
        photo: "Coil group on winding form showing turns count verification.",
        accept: "Turns, wire gauge, strands, and coil geometry per winding specification.", record: "Wire lot, turns per coil, winder ID, operator." },
      { op: "30", dept: "WIND", title: "Insert Coils per Work Instruction / Drawing",
        text: `Install ${ref(/insulation/i, "slot insulation")} and insert coil groups into ${ref(/lamination stack/i, "the lamination stack")} per work instruction and drawing: slot fill sequence, phase placement, and pitch per released data. Install phase insulation and slot wedges.`,
        sub: ["Install slot liners per drawing; verify no liner damage.", "Insert coils in released sequence and pitch — no forced insertion, no lamination or wire insulation damage.", "Install phase-to-phase insulation and top sticks / wedges per drawing.", "Verify bore is clear of protruding insulation and wire."],
        photo: "Coil insertion showing slot fill, phase insulation, and wedge installation.",
        accept: "Placement, pitch, and slot fill per drawing; no insulation or lamination damage.", record: "Insertion sequence verification, operator." },
      { op: "40", dept: "TEST", title: "Verify Electrical Rotation",
        text: "Verify coil group polarity and phase sequence (electrical rotation) per released test method before lead connection. Correct any group polarity errors prior to proceeding.",
        accept: "Electrical rotation / phase sequence correct per drawing.", record: "Rotation check result, method/equipment ID, operator." },
      { op: "50", dept: "ASSEMBLY", title: "Connect Lead Wires",
        text: `Make phase and neutral connections and attach ${ref(/leadwire|lead wire/i, "leadwire")} per drawing: joint method per released process, sleeve and insulate joints, apply phase identification, and secure with released strain relief.`,
        sub: [`Verify ${refPN(/leadwire|lead wire/i)} type, gauge, and length per drawing.`, "Make joints per released process (weld / crimp / solder as specified).", "Sleeve and insulate each joint per workmanship standard.", "Apply phase ID markers; route and secure leads with strain relief."],
        photo: "Completed phase joints sleeved with phase identification visible.",
        accept: "Joint workmanship, phase ID, routing, and strain relief per drawing.", record: "Lead lot, joint method, operator." },
      { op: "60", dept: "ASSEMBLY", title: "Lace / Finish Coils",
        text: `Lace end turns with ${ref(/lacing/i, "lacing cord")} per workmanship standard: secure coil extensions and lead dress, maintain phase separation, no loose wires.`,
        accept: "Lacing complete and uniform; leads secured; phase separation maintained.", record: "Lacing material lot, operator." },
      { op: "70", dept: "ASSEMBLY", title: "Form End Turns",
        text: "Form end turns to drawing envelope using released forming method/tooling. Verify bore clearance with released gauge and end-turn height/OD per drawing. No conductor or insulation damage.",
        photo: "Formed end turns with envelope gauge in place.",
        accept: "End-turn envelope, bore clearance, and height per drawing; no damage.", record: "Gauge verification, forming tool ID, operator." },
      { op: "80", dept: "TEST", title: "Pre-Varnish Electrical Testing", hold: true,
        text: "Perform pre-impregnation electrical tests per released values: phase resistance and balance, insulation resistance, surge comparison, dielectric withstand.",
        accept: "All pre-varnish electrical results acceptable and recorded. QA accepts before impregnation.", record: "All test values, equipment IDs, tester stamp, QA stamp." },
      { op: "90", dept: "PROC", title: "Varnish / Impregnation of Coils",
        text: `Impregnate the wound stator with ${ref(/varnish|resin/i, "approved varnish")} per released vacuum/pressure or dip-and-bake process. Cure strictly per approved varnish specification; record ${refPN(/varnish|resin/i)} lot and full cycle data.`,
        accept: "Impregnation and cure per released process; no uncured varnish or blocked bore.", record: "Varnish lot, cycle data, drain time, oven ID, cure chart." },
      { op: "100", dept: "MFG", title: "Clean Up / Visual Inspection (Manufacturing)",
        text: "Remove excess varnish from bore, mounting surfaces, and leads per released method. Manufacturing visual inspection: fill quality, no voids or runs, no loose wires or sharp varnish, lead condition acceptable.",
        accept: "Bore and mounting surfaces clean; visual workmanship acceptable to manufacturing.", record: "Cleanup method, mfg inspection result, operator." },
      { op: "110", dept: "TEST", title: "Final Test", hold: true,
        text: "Perform final electrical acceptance test per released ATP: resistance/balance, insulation resistance, surge, dielectric withstand, inductance, and electrical rotation verification. Retain electronic record.",
        accept: "All values meet released ATP.", record: "ATP data file, equipment IDs, tester stamp." },
      { op: "120", dept: "QA / INSPECT", title: "Final QA Inspection", hold: true,
        text: "Final dimensional and visual inspection: bore, mounting interfaces, end-turn envelope, lead condition and identification, workmanship, cleanliness. Verify traveler completeness and, for matched sets, pairing record with mating rotor serial.",
        accept: "All characteristics accepted; traveler complete; QA final acceptance.", record: "Inspection report, final QA stamp/date." },
      { op: "130", dept: "STOCK", title: "Stock for NHA",
        text: "Preserve and protect leads and bore; package with contamination protection; identify with lot/serial; complete stock transaction and release to next higher assembly.",
        accept: "Preservation, identification, and stock transaction per released requirements.", record: "Qty, stock location, lot/serial, NHA reference." },
    ] },
  { id: "statorStack", match: p => /lamination|stator core|stack\s+ass|stator/i.test(p.desc) && !/winding|coil/i.test(p.desc),
    atp: [["Stack height / squareness", "Per drawing"], ["Bore / OD size", "Per drawing"], ["Interlaminar insulation", "Per released test"]],
    safety: ["epoxy"],
    ops: (p, kids, bom, excluded, ref, refPN) => [
      { op: "10", dept: "STORES / QA", title: "Kitting", hold: true, text: `Kit ${ref(/lamination/i, "laminations")} and ${ref(/epoxy|adhesive/i, "bonding epoxy")} plus consumables per ${p.pn} BOM. Verify lamination lot, coating/insulation condition, and adhesive shelf life.`, accept: "BOM and kitting form compliant; lots recorded.", record: "Kitting form, lots, certs, QA stamp." },
      { op: "20", dept: "QA / INSPECT", title: "Lamination Inspection", text: `Sample-inspect ${ref(/lamination/i, "laminations")} per released plan: burrs, coating damage, flatness, and dimensional conformance.`, accept: "Sample accepted per plan; no disqualifying damage.", record: "Sample results, inspector stamp." },
      { op: "30", dept: "ASSEMBLY", title: "Stack & Bond", text: `Stack ${refPN(/lamination/i)} laminations to count/height in the stacking fixture with released orientation/rotation pattern. Bond with ${ref(/epoxy|adhesive/i, "bonding epoxy")} per released process.`,
        sub: ["Verify fixture ID and condition.", `Stack ${refPN(/lamination/i)} to released count; apply rotation/interleave pattern per drawing.`, `Apply ${refPN(/epoxy|adhesive/i)} per released method; maintain alignment of slots/bore.`, "Clamp to released pressure; verify stack height and squareness before cure."],
        photo: "Lamination stacking fixture with slot alignment feature engaged.",
        accept: "Count/height and alignment per drawing before cure.", record: "Count, height, adhesive lot/mix, fixture ID." },
      { op: "40", dept: "PROC / QA", title: "Cure & Post-Cure Inspection", hold: true, text: "Cure per released schedule; record cure data. Inspect for delamination, resin voids, slot obstruction, and dimensional conformance.", accept: "Cure within schedule; stack conforms; slots clear.", record: "Cure start/stop, oven ID, inspection results, QA stamp." },
      { op: "50", dept: "MACHINE", title: "Finish Machining (as required)", text: "Machine/grind bore, OD, and faces to drawing as applicable. Deburr; protect insulation; clean thoroughly.", accept: "Dimensions per drawing; no insulation damage.", record: "Dimensions, equipment ID." },
      { op: "60", dept: "QA / INSPECT", title: "Final Inspection", hold: true, text: "Final dimensional/visual inspection and interlaminar insulation check per released test.", accept: "All characteristics accepted.", record: "Inspection report, final QA stamp." },
      { op: "70", dept: "STOCK", title: "Preserve & Stock", text: "Clean, preserve, identify, and stock for winding / next higher assembly.", accept: "Preservation and ID per requirements.", record: "Qty, location, lot ID." },
    ] },
  { id: "motor", match: p => /motor.{0,28}(assembly|assy)|(assembly|assy).{0,28}motor/i.test(p.desc),
    atp: [["Phase resistance U-V / V-W / W-U", "Per ATP at stated temp"], ["Resistance imbalance", "Per ATP"], ["Insulation resistance", "Per ATP"], ["Dielectric withstand", "No breakdown"], ["Back-EMF / Ke", "Per ATP at controlled RPM"], ["No-load current / speed", "Per ATP"], ["Axial endplay / preload", "Per drawing"], ["Direction / commutation", "Correct"]],
    safety: ["magnets", "electrical"],
    ops: (p, kids, bom, excluded, ref, refPN) => {
      const has = re => kids.some(k => !excluded[k.pn] && re.test((bom.parts[k.pn] || {}).desc || ""));
      const o = [
        { op: "10", dept: "KIT", title: "Kitting", hold: true,
          text: `Kit all components per ${p.pn} BOM, including the completed and accepted ${ref(/stator.*assembl/i, "Stator Winding Assembly")} and ${ref(/rotor assembly/i, "Rotor Assembly")} (serialized), ${ref(/bearing/i, "bearings")}, ${ref(/housing/i, "housing")}, ${ref(/front endbell/i, "front endbell")}, ${ref(/rear endbell/i, "rear endbell")}, and hardware. Verify subassembly traveler completion and lot traceability.`,
          sub: [`Verify ${refPN(/stator.*assembl/i)} and ${refPN(/rotor assembly/i)} travelers are complete and QA-accepted.`, "Record serialized subassembly identities and matched-set references.", `Verify ${refPN(/bearing/i)} matched-set lot and approved source.`, "Complete the kitting form and attach certifications."],
          accept: "QA kitting verification complete; subassembly acceptance verified; lots and serials recorded.", record: "Kitting form, subassembly serials, bearing lots." },
        { op: "20", dept: "ASSEMBLY", title: "Inspect & Clean Interfaces",
          text: "Inspect housing bore, stator OD, bearing seats, endbell pilots, and shaft journals. Clean with approved solvent; protect windings, leads, and insulation.",
          accept: "Surfaces clean, dry, damage-free, within drawing.", record: "Cleaning method; solvent lot if required." },
        { op: "30", dept: "ASSEMBLY", title: "Install Stator Winding Assembly into Housing",
          text: `Install the completed ${ref(/stator.*assembl/i, "stator winding assembly")} into ${ref(/housing/i, "the motor housing")} per released process (thermal shrink, press, or bond as specified). Orient leads per drawing; protect windings and end turns throughout installation.`,
          sub: [`Verify ${refPN(/housing/i)} bore and ${refPN(/stator.*assembl/i)} OD dimensions/fit class per drawing.`, "Heat housing / press / apply bonding agent strictly per released process.", "Orient stator so lead exit matches drawing clocking.", "Verify full axial seating against the released datum.", "Allow controlled cooldown / cure before next operation per released process."],
          photo: "Stator installation showing lead clocking orientation and seating verification.",
          callout: { k: "CAUTION", v: "Do not apply installation force through windings or end turns. Damage to wire insulation during installation is cause for rejection." },
          accept: "Stator fully seated at drawing position; lead orientation correct; no winding damage.", record: "Fit method, temperatures/forces per released process, operator." },
        { op: "40", dept: "TEST", title: "Post-Installation Electrical Verification", hold: true,
          text: "Verify stator electricals after installation: phase resistance and insulation resistance within released limits relative to the stator final test record. Detects installation damage before rotor installation.",
          accept: "Values within released delta of stator final-test record. QA accepts before rotor installation.", record: "Measured values vs. stator record, equipment ID, QA stamp." },
      ];
      let n = 50;
      o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Rotor & Bearings",
        text: `Install ${ref(/bearing/i, "bearings")} with controlled force applied only to the ring being fitted — approved arbor tooling or controlled heating; never through rolling elements. Guide ${ref(/rotor assembly/i, "the rotor")} into the stator preventing magnet-to-stator impact.`,
        sub: [`Verify ${refPN(/bearing/i)} lot and matched-set ID.`, `Press/heat-fit ${refPN(/bearing/i)} per released method; force on fitted ring only.`, `Use guided insertion tooling for ${refPN(/rotor assembly/i)} — magnets pull toward stator; maintain control at all times.`, "Verify rotor turns freely with no scraping."],
        photo: "Guided rotor insertion tooling controlling magnetic pull-in.",
        callout: { k: "WARNING", v: "Magnetic pull-in force between rotor and stator is large and sudden. Use insertion tooling — never insert by hand." },
        accept: "Bearings fully seated; rotor free; no brinelling or impact damage.", record: "Bearing lots, method, operator." }); n += 10;
      o.push({ op: String(n), dept: "ASSEMBLY", title: "Close & Set Endplay",
        text: `Install ${ref(/front endbell/i, "front endbell")}, ${ref(/rear endbell/i, "rear endbell")}, ${ref(/wave spring/i, "wave spring")}/shims, retainers, and fasteners${has(/hardware kit/i) ? ` from ${refPN(/hardware kit/i)} hardware kit` : ""}. Set axial endplay/preload per drawing. Torque in sequence and witness-mark.`,
        accept: "Endplay/preload, runout, torque per drawing.", record: "Endplay/preload, shim stack, torques, wrench ID." }); n += 10;
      if (has(/connector/i)) { o.push({ op: String(n), dept: "ASSEMBLY", title: "Terminate Leads to Connector",
        text: `Terminate stator leads to ${ref(/connector/i, "the connector")} per drawing: contact crimp/solder per released process, pinout per drawing, strain relief and sealing as specified. Verify pinout continuity.`,
        accept: "Pinout, workmanship, and continuity per drawing.", record: "Continuity results, contact lot, operator." }); n += 10; }
      o.push({ op: String(n), dept: "TEST", title: "Motor Acceptance Test", hold: true,
        text: "Full ATP: resistance/balance, IR, dielectric, back-EMF/Ke, commutation phasing, no-load current/speed/direction, vibration/noise, sensor verification.",
        accept: "All values meet released ATP; electronic record retained.", record: "ATP data file, equipment IDs, tester stamp." }); n += 10;
      o.push({ op: String(n), dept: "QA", title: "Final Inspection", hold: true,
        text: "Configuration, lead/connector ID, pinout, shaft condition, witness marks, cleanliness, nameplate, serialization. Verify all subassembly and motor travelers complete.",
        accept: "Final QA acceptance.", record: "Final QA stamp/date." }); n += 10;
      o.push({ op: String(n), dept: "STOCK", title: "Preserve & Stock",
        text: "Cap connector/shaft; package with corrosion and ESD protection; stock for next higher assembly or shipment.",
        accept: "Stock transaction complete.", record: "Location, qty, serial." });
      return o;
    } },
  { id: "brake", match: p => /brake assembly/i.test(p.desc),
    atp: [["Coil resistance", "Per drawing"], ["Release / dropout voltage or current", "Per ATP"], ["Air gap", "Per drawing"], ["Holding torque", "Per ATP"]],
    safety: ["electrical"],
    ops: (p, kids, bom, excluded, ref, refPN) => [
      { op: "10", dept: "STORES / QA", title: "Kitting", hold: true, text: `Kit ${ref(/coil/i, "coil")}, ${ref(/back iron/i, "back iron")}, ${ref(/armature/i, "armature")}, ${ref(/spring/i, "springs")}, ${ref(/friction/i, "friction disc")}, and hardware per ${p.pn} BOM. Verify plating condition and traceability.`, accept: "BOM/kitting compliant.", record: "Kitting form, lots." },
      { op: "20", dept: "ASSEMBLY", title: "Coil Installation", text: "Install/pot brake coil in back iron per released process. Verify lead routing and strain relief.", accept: "Coil seated; leads per drawing.", record: "Coil lot, potting lot if applicable." },
      { op: "30", dept: "TEST", title: "Coil Electrical Check", text: "Verify coil resistance and insulation resistance per released values.", accept: "Values per drawing/ATP.", record: "Measured values, equipment ID." },
      { op: "40", dept: "ASSEMBLY", title: "Armature, Springs & Friction Disc", text: "Install compression springs, armature plate, and friction disc per drawing orientation.",
        sub: ["Verify spring free length (sample) per drawing.", "Install springs in released pattern.", "Install armature — verify flatness/orientation.", "Install friction disc; verify no contamination on friction surfaces."],
        photo: "Spring pattern and armature orientation before closing.",
        accept: "Stack-up per drawing; friction surfaces clean/dry.", record: "Spring lot, disc lot." },
      { op: "50", dept: "ASSEMBLY", title: "Set Air Gap", hold: true, text: "Set working air gap per released shim/setting procedure. Torque fasteners and apply witness marks.", accept: "Air gap within drawing limit.", record: "Air gap, shim stack, torques." },
      { op: "60", dept: "TEST", title: "Functional Test", hold: true, text: "Verify release/dropout at released values; verify holding function per ATP.", accept: "All ATP values met.", record: "Test values, equipment ID." },
      { op: "70", dept: "QA / STOCK", title: "Final Inspect & Stock", text: "Final visual/config inspection; preserve, identify, and stock.", accept: "Final QA acceptance.", record: "Qty, location, QA stamp." },
    ] },
  { id: "gearhead", match: p => /gear[\s-]*head/i.test(p.desc),
    atp: [["Backlash / lost motion", "Per drawing"], ["Running torque", "Per ATP"], ["Output runout", "Per drawing"]],
    safety: [],
    ops: (p, kids, bom, excluded, ref, refPN) => [
      { op: "10", dept: "STORES / QA", title: "Kitting", hold: true, text: `Kit ${ref(/gear housing/i, "housing")}, ${ref(/sun gear/i, "sun gear")}, ${ref(/planet gear/i, "planet gears")}, ${ref(/carrier/i, "carrier")}, ${ref(/ring gear/i, "ring gear")}, ${ref(/bearing/i, "bearings")}, and hardware per ${p.pn} BOM.`, accept: "BOM/kitting compliant.", record: "Kitting form, lots." },
      { op: "20", dept: "QA / INSPECT", title: "Gear & Bearing Inspection", text: "Inspect gear teeth (nicks, burrs, heat treat stamp), bearing lots, and housing bores per drawing.", accept: "No disqualifying damage; dims per drawing.", record: "Results, inspector stamp." },
      { op: "30", dept: "ASSEMBLY", title: "Press Bearings", text: "Press bearings into housing/carrier with approved tooling; force on fitted ring only.", accept: "Fully seated; no brinelling.", record: "Bearing lots, method." },
      { op: "40", dept: "ASSEMBLY", title: "Assemble Gear Train", text: `Install ${ref(/sun gear/i, "sun gear")}, ${ref(/planet gear/i, "planet gears")} with ${ref(/carrier/i, "carrier")}, and ${ref(/ring gear/i, "ring gear")} per drawing. Verify timing marks/phasing where applicable. Apply released lubricant.`,
        sub: ["Verify planet gear set is matched where required.", "Assemble with released lubricant type and quantity.", "Verify free rotation through full revolution — no tight spots."],
        photo: "Planet carrier assembly showing gear phasing/timing marks.",
        accept: "Free rotation; phasing correct; lube per released spec.", record: "Lube type/lot/qty, operator." },
      { op: "50", dept: "TEST", title: "Backlash & Running Torque", hold: true, text: "Measure backlash/lost motion and running torque per released method.", accept: "Values per drawing/ATP.", record: "Measured values, equipment ID." },
      { op: "60", dept: "QA / STOCK", title: "Final Inspect & Stock", text: "Final inspection; preserve, identify, stock for next higher assembly.", accept: "Final QA acceptance.", record: "Qty, location, QA stamp." },
    ] },
  { id: "actuator", match: p => /actuator/i.test(p.desc),
    atp: [["Motor phase-to-phase resistance", "Per ATP; balanced"], ["Insulation resistance", "Per ATP"], ["Dielectric withstand", "No breakdown"], ["No-load current / speed", "Per ATP"], ["Gearhead backlash / lost motion", "Per drawing/ATP"], ["Brake release current", "Per ATP"], ["Brake dropout / re-engagement", "Per ATP"], ["Brake holding torque", "Per ATP"], ["Output runout / endplay", "Per drawing"]],
    safety: ["electrical"],
    ops: (p, kids, bom, excluded, ref, refPN) => {
      const has = re => kids.some(k => !excluded[k.pn] && re.test((bom.parts[k.pn] || {}).desc || ""));
      const o = [
        { op: "10", dept: "KIT", title: "Kitting", hold: true, text: `Kit all released components and completed lower-level assemblies per ${p.pn} BOM${has(/gearhead/i) ? `, including ${ref(/gearhead/i)}` : ""}${has(/motor assembly/i) ? `, ${ref(/motor assembly/i)}` : ""}${has(/brake/i) ? `, ${ref(/brake/i)}` : ""}${has(/encoder|sensor/i) ? `, ${ref(/encoder|sensor/i)}` : ""}. Verify PN, revision, qty, shelf life, lot traceability, cert status, preservation.`, accept: "QA verifies BOM/kitting compliance; serialized subassemblies recorded.", record: "Kitting form, serial IDs." },
        { op: "20", dept: "ASSEMBLY", title: "Clean & Inspect Interfaces", text: "Clean and inspect mating faces, pilots, dowels, threads, electrical interfaces, seals, connector surfaces. No burrs, corrosion, damage, FOD.", accept: "Visual inspection acceptable.", record: "Cleaning method; solvent lot if required." },
      ];
      let n = 30;
      if (has(/gearhead/i)) { o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Gearhead", text: `Install ${ref(/gearhead/i, "gearhead")} to ${ref(/main housing|housing/i, "main housing")}. Engage pilot/dowels without forcing. Approved threadlocker where specified. Torque cross-pattern per released table.`, photo: "Cross-pattern torque sequence on gearhead-to-housing fasteners.", accept: "Full seating, no visible gap; torque recorded.", record: "Torque wrench ID, final torque." }); n += 10; }
      o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Motor Assembly", text: `Install ${ref(/motor assembly/i, "motor")} and verify shaft/coupling engagement. Do not transmit assembly force through motor bearings. Verify axial seating and free rotation before tightening.`, accept: "Shaft rotates smoothly by hand; no binding or abnormal endplay.", record: "Operator, torque data." }); n += 10;
      if (has(/brake/i)) { o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Brake & Set Gap", text: `Install ${ref(/brake/i, "brake")} and hub/coupling. Set working air gap or axial location per released shim procedure. Torque and witness-mark.`, accept: "Air gap within drawing limit.", record: "Air gap, shim stack, torques." }); n += 10; }
      if (has(/encoder|sensor|resolver/i)) { o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Encoder / Feedback", text: `Install ${ref(/encoder|sensor|resolver/i, "feedback device")}; align zero/index per drawing or electrical setup. Route harness with released bend radius, clearance, strain relief, chafe protection.`, photo: "Encoder zero/index alignment setup.", accept: "Alignment recorded; harness retained; continuity verified.", record: "Alignment value, continuity results." }); n += 10; }
      if (has(/harness|cable/i)) { o.push({ op: String(n), dept: "ASSEMBLY", title: "Install Harness / Cable", text: `Install and route ${ref(/harness|cable/i, "harness")} per drawing: bend radius, strain relief, clocking, connector torque. Verify pinout continuity.`, accept: "Routing and continuity per drawing.", record: "Continuity results, connector torque." }); n += 10; }
      o.push({ op: String(n), dept: "TEST", title: "Low-Speed Functional Test", hold: true, text: "Energize with current-limited source. Verify rotation direction, commutation, brake release/re-engagement, absence of abnormal noise/vibration.", accept: "Direction correct; values within released ATP.", record: "Current, speed, brake values." }); n += 10;
      o.push({ op: String(n), dept: "TEST", title: "Acceptance Test (ATP)", hold: true, text: "Full ATP under representative load where fixture available: currents, speed, backlash, running torque, brake holding torque.", accept: "All ATP results acceptable.", record: "ATP data file, equipment IDs." }); n += 10;
      o.push({ op: String(n), dept: "QA", title: "Final Inspection", hold: true, text: "Final visual/dimensional: nameplate/serialization, witness marks, connector ID, workmanship, cleanliness, configuration vs. released BOM, closure of all prior ops.", accept: "Final inspection accepted; configuration matches released BOM.", record: "Final QA stamp/date." }); n += 10;
      o.push({ op: String(n), dept: "STOCK", title: "Preserve, Package & Ship/Stock", text: "Protective caps and preservation; package per approved instruction; transfer to controlled stock or shipment.", accept: "Stock transaction complete.", record: "Location, qty." });
      return o;
    } },
  { id: "generic", match: () => true, atp: null, safety: [],
    ops: (p) => [
      { op: "10", dept: "STORES / QA", title: "Kitting", hold: true, text: `Kit all components per ${p.pn} BOM. Verify PN, rev, qty, shelf life, and traceability.`, accept: "BOM/kitting compliant.", record: "Kitting form, lots." },
      { op: "20", dept: "QA / INSPECT", title: "Pre-Assembly Inspection", text: "Inspect components per drawing; verify no damage or contamination.", accept: "Conforms to drawings.", record: "Results, inspector stamp." },
      { op: "30", dept: "ASSEMBLY", title: "Assemble", text: `Assemble ${p.desc || p.pn} per released drawing and process. Torque fasteners per released table; witness-mark.`, accept: "Assembly per drawing; torques recorded.", record: "Torques, operator." },
      { op: "40", dept: "QA", title: "Final Inspection", hold: true, text: "Final dimensional/visual inspection per drawing.", accept: "All characteristics accepted.", record: "Final QA stamp/date." },
      { op: "50", dept: "STOCK", title: "Preserve & Stock", text: "Preserve, identify, and stock for next higher assembly.", accept: "Stock transaction complete.", record: "Qty, location." },
    ] },
];

/* ---------------- Scope helpers ---------------- */
function scopedTree(bom, excluded, top) {
  const out = [];
  const isAsm = pn => !!(bom.children[pn] && bom.children[pn].length);
  function walk(pn, depth, row) {
    if (excluded[pn]) return;
    const part = bom.parts[pn] || { pn, desc: "(undefined)", rev: "-", mb: "" };
    out.push({ depth, row, part, isAsm: isAsm(pn) });
    for (const k of (bom.children[pn] || [])) walk(k.pn, depth + 1, k);
  }
  walk(top, 0, null);
  return out;
}
function buildOrder(bom, excluded, tops) {
  const seen = {}, order = [];
  const isAsm = pn => !!(bom.children[pn] && bom.children[pn].length);
  function walk(pn) {
    if (excluded[pn] || seen[pn]) return;
    seen[pn] = 1;
    for (const k of (bom.children[pn] || [])) walk(k.pn);
    const p = bom.parts[pn];
    if (isAsm(pn) && p && /make/i.test(p.mb || "Make")) order.push(pn);
  }
  tops.forEach(walk);
  return order;
}
function opsFor(bom, excluded, pn, misses) {
  const p = bom.parts[pn], kids = bom.children[pn] || [];
  const tpl = activeLibrary().find(t => t.match(p));
  // resolve actual BOM child part numbers into op text at generation time
  const find = re => kids.find(k => !excluded[k.pn] && re.test((bom.parts[k.pn] || {}).desc || ""));
  const ref = (re, fb) => { const k = find(re); if (!k) { if (misses) misses.push({ assembly: pn, kind: "ref", pattern: re.source, fallback: fb || "per BOM" }); return fb || "per BOM"; }
    const pp = bom.parts[k.pn];
    return `${pp.desc} (${k.pn}${k.qty && k.qty !== "1" ? `, ${k.qty} ${k.uom}` : ""})`; };
  const refPN = (re, fb) => { const k = find(re); if (!k && misses) misses.push({ assembly: pn, kind: "refPN", pattern: re.source, fallback: fb || "______" }); return k ? k.pn : (fb || "______"); };
  return { tpl, ops: tpl.ops(p, kids, bom, excluded, ref, refPN) };
}

/* =====================================================================
   TEMPLATE JSON LAYER — default generation logic unchanged; imported
   declarative JSON templates override by id / take precedence.
   Text tokens: {PN} {DESC} {ref:regex|fallback} {refPN:regex|fallback}
   Per-op: onlyIf (regex vs child descriptions). autoNumber renumbers *10.
   ===================================================================== */
let CUSTOM_TEMPLATES = []; // interpreted, checked before OPS_LIBRARY

function makeOpsFromDecl(decl) {
  return (p, kids, bom, excluded, ref, refPN) => {
    const has = re => kids.some(k => !excluded[k.pn] && re.test((bom.parts[k.pn] || {}).desc || ""));
    let ops = (decl.ops || []).filter(o => !o.onlyIf || has(new RegExp(o.onlyIf, "i")));
    if (decl.autoNumber) ops = ops.map((o, i) => ({ ...o, op: String((i + 1) * 10) }));
    const sub = t => t == null ? t : String(t)
      .replace(/\{PN\}/g, p.pn).replace(/\{DESC\}/g, p.desc || p.pn)
      .replace(/\{ref:([^|}]+)\|?([^}]*)\}/g, (_, r, fb) => ref(new RegExp(r, "i"), fb || undefined))
      .replace(/\{refPN:([^|}]+)\|?([^}]*)\}/g, (_, r, fb) => refPN(new RegExp(r, "i"), fb || undefined));
    return ops.map(o => ({ ...o, title: sub(o.title), text: sub(o.text), accept: sub(o.accept), record: sub(o.record),
      photo: o.photo ? sub(o.photo) : o.photo, sub: o.sub ? o.sub.map(sub) : o.sub,
      callout: o.callout ? { k: o.callout.k, v: sub(o.callout.v) } : o.callout }));
  };
}
function setCustomTemplates(decls) {
  CUSTOM_TEMPLATES = (decls || []).map(d => ({
    id: d.id, custom: true,
    match: p => { try { return new RegExp(d.match, "i").test(p.desc || ""); } catch (e) { return false; } },
    atp: d.atp || null, safety: d.safety || [], ops: makeOpsFromDecl(d),
  }));
}
function activeLibrary() {
  // custom templates first; a custom id shadows the built-in of the same id
  const shadowed = new Set(CUSTOM_TEMPLATES.map(t => t.id));
  return [...CUSTOM_TEMPLATES, ...OPS_LIBRARY.filter(t => !shadowed.has(t.id))];
}
function exportTemplatesJSON() {
  // token-preserving stub execution: {ref:...} tokens survive into exported text
  const stubP = { pn: "{PN}", desc: "{DESC}", mb: "Make" };
  const allKeywords = "gearhead motor assembly brake encoder sensor resolver harness cable connector hardware kit stator rotor lamination stack winding housing endbell wave spring bearing";
  const stubKids = [{ pn: "{ANY}", qty: "1", uom: "EA" }];
  const stubBom = { parts: { "{ANY}": { pn: "{ANY}", desc: allKeywords } }, children: {} };
  const tRef = (re, fb) => `{ref:${re.source}|${fb || ""}}`;
  const tRefPN = (re, fb) => `{refPN:${re.source}|${fb || ""}}`;
  return OPS_LIBRARY.map(t => {
    let ops = [];
    try { ops = t.ops(stubP, stubKids, stubBom, {}, tRef, tRefPN); } catch (e) { ops = []; }
    const ms = t.match.toString();
    const firstRe = (ms.match(/\/((?:[^\/\\]|\\.)+)\/[a-z]*/) || [])[1] || t.id;
    return { id: t.id, match: firstRe, matchSource: ms, atp: t.atp || null, safety: t.safety || [],
      ops: ops.map(o => ({ op: o.op, dept: o.dept, title: o.title, hold: !!o.hold, text: o.text,
        sub: o.sub, accept: o.accept, record: o.record, photo: o.photo, callout: o.callout })) };
  });
}

/* =====================================================================
   DOCUMENT CHECK — in-app version of the dev validation harness
   ===================================================================== */
function runDocumentCheck(bom, excluded, tops, purchased) {
  const errors = [], warns = [], info = [];
  (bom.issues || []).forEach(i => warns.push("BOM: " + i));
  const misses = [];
  const order = buildOrder(bom, excluded, tops);
  if (!order.length) errors.push("No Make assemblies in scope — nothing to route.");
  for (const pn of order) {
    const { tpl, ops } = opsFor(bom, excluded, pn, misses);
    const nums = ops.map(o => +o.op);
    if (!nums.every((n, i) => i === 0 || n > nums[i - 1])) errors.push(`${pn}: op numbers not strictly ascending.`);
    if (!ops.some(o => o.hold)) errors.push(`${pn}: no QA hold point in routing.`);
    if (!/kit/i.test(ops[0].title)) warns.push(`${pn}: first op is not kitting (${ops[0].title}).`);
    if (!/stock|ship/i.test(ops[ops.length - 1].title)) warns.push(`${pn}: last op is not stock/ship (${ops[ops.length - 1].title}).`);
    if (tpl.id === "generic") info.push(`${pn}: routed with the generic template — consider a specific template for "${(bom.parts[pn] || {}).desc}".`);
    if (tpl.custom) info.push(`${pn}: routed with imported custom template "${tpl.id}".`);
  }
  Object.values(bom.parts).forEach(p => {
    if (!(bom.children[p.pn] && bom.children[p.pn].length) && isAssemblyLike(p) && !(purchased || {})[p.pn])
      warns.push(`${p.pn} ("${p.desc}") has no BOM loaded — drop the sub-BOM file or mark purchased.`);
  });
  const seen = new Set();
  misses.forEach(m => {
    const k = m.assembly + "|" + m.pattern;
    if (seen.has(k)) return; seen.add(k);
    info.push(`${m.assembly}: unresolved BOM reference /${m.pattern}/ — using fallback "${m.fallback}". Tune the description matcher or BOM wording.`);
  });
  return { errors, warns, info, pass: errors.length === 0 };
}

/* ===== ERP BOM report parsers (JobBoss / ICG-style exports) =====
   Formats:
   A. Exploded BOM (level notation "2******") — full tree, level mode
   B. Standard BOM (sections "Part No:" + Assy rows) — parent mode, multi-section
   C. Report grids from xls (same reports, plus FRM-61 kitting form) — column-true
   All output: { mode:'level'|'parent', rows:[{level?,parent?,find,qty,uom,pn,rev,desc,mb,mat,rem}], meta } */

const RPT_UNITS = "EA|QT|OZ|FT|YD|LB|GA|L|IN2|IN|PC|SET|KT|EA\\.";
const RPT_PN_RE = /^[A-Z0-9][A-Z0-9./#*-]{2,}$/;
const isPN = t => !!t && RPT_PN_RE.test(t) && /\d/.test(t) && !new RegExp("^(" + RPT_UNITS + ")$").test(t) && !/^\d[\d,.]*$/.test(t) && !/^\$\d/.test(t);
const SKIP_RE = /^(Page \d+ of|Report Generated|Island Components|Bill of Materials|Exploded Bill|Level\s+Part Number|Stock\s+Quantity|Part\s+Unit$|Type\s+Number|Description Required Net|Quantity$|TOTAL MATERIAL|Part No:|Alt Part No:|Used On:|Description:|Descrip:|Cautions?:|~~~|\d+-\s|For Assy|\$[\d,.]+(\s+\$[\d,.]+)?$|Revision:|Quantity:\s|Enter QTY|Prepared By|Received By|QC Verify)/i;
const numQty = s => { const n = parseFloat(String(s).replace(/,/g, "")); return isNaN(n) ? "1" : (n === Math.floor(n) ? String(n) : String(n)); };

function detectReportFormat(text) {
  const t = String(text || "");
  const levelLines = (t.match(/^\s*\d\*{2,}/gm) || []).length;
  if (/Exploded Bill of Material/i.test(t) || levelLines >= 3) return "exploded";
  if (/Bill of Materials/i.test(t) && /Part No:/i.test(t)) return "standard";
  return null;
}

/* ---- A. Exploded BOM text (also FRM-61-style level text) ---- */
function parseExplodedText(text) {
  const lines = String(text).split(/\r?\n/).map(l => l.trim());
  const rows = []; const notes = [];
  let pendingLevel = null, cur = null, curDone = false;
  const dataTail = new RegExp("^(.*?)\\s{2,}(" + RPT_UNITS + ")\\s+([\\d,]+\\.?\\d*)\\s+\\$");
  const dataTailLoose = new RegExp("^(.*?)\\s+(" + RPT_UNITS + ")\\s+([\\d,]+\\.?\\d*)\\s+\\$");
  const unitFirst = new RegExp("^(" + RPT_UNITS + ")\\s+([\\d,]+\\.?\\d*)\\s+\\$"); // chat-copy variant: data on its own line
  const flush = () => { if (cur) { cur.desc = cur.desc.replace(/\s+/g, " ").trim(); rows.push(cur); } cur = null; curDone = false; };
  for (const raw of lines) {
    if (!raw) continue;
    let l = raw;
    const lvlOnly = l.match(/^(\d)\*{2,}$/);
    if (lvlOnly) { flush(); pendingLevel = +lvlOnly[1]; continue; }
    const lvlLead = l.match(/^(\d)\*{2,}\s+(.*)$/);
    let level = null;
    if (lvlLead) { flush(); level = +lvlLead[1]; l = lvlLead[2]; }
    else if (pendingLevel != null && isPN(l.split(/\s+/)[0])) { level = pendingLevel; }
    if (level != null) {
      pendingLevel = null;
      const toks = l.split(/\s+/);
      const pn = toks[0];
      if (!isPN(pn)) { continue; }
      let rest = l.slice(pn.length).trim();
      cur = { level, parent: null, find: "", qty: "1", uom: "EA", pn, rev: "-", desc: "", mb: "", mat: "", rem: "" };
      const m = rest.match(dataTail) || rest.match(dataTailLoose);
      if (m) { cur.desc = m[1]; cur.uom = m[2]; cur.qty = numQty(m[3]); curDone = true; }
      else cur.desc = rest;
      continue;
    }
    if (SKIP_RE.test(l)) { flush(); continue; }
    if (cur) {
      if (!curDone) {
        const m2 = l.match(unitFirst) || l.match(dataTailLoose);
        if (m2 && m2.length === 4) { cur.desc += " " + m2[1]; cur.uom = m2[2]; cur.qty = numQty(m2[3]); curDone = true; continue; }
        if (m2) { cur.uom = m2[1]; cur.qty = numQty(m2[2]); curDone = true; continue; }
      }
      // PN wrap rejoin: a leading 1-2 char token on the first continuation line completes the part number
      if (curDone && !cur._pnFixed) {
        const pw = l.match(/^([A-Z0-9]{1,2})(?:$|\s{2,}(.*))/);
        if (pw) { cur.pn += pw[1]; cur._pnFixed = true; if (pw[2]) cur.desc += " " + pw[2]; continue; }
        cur._pnFixed = true; // only the first continuation line is eligible
      }
      if (!/^\$|^[\d,\s.$-]+$/.test(l)) cur.desc += " " + l;  // continuation (skip pure-number columns)
    }
  }
  flush();
  rows.forEach(r => delete r._pnFixed);
  return { mode: "level", rows, meta: { format: "exploded", notes } };
}

/* ---- B. Standard BOM text (sectioned) ---- */
function parseStandardText(text) {
  const lines = String(text).split(/\r?\n/).map(l => l.trim());
  const rows = []; const sections = []; const notes = [];
  let parent = null, parentRev = "-", expectPN = false, cur = null;
  const rowRe = new RegExp("^(?:Assy\\s+)?(\\S+)\\s+(.*?)\\s+([\\d,]+\\.?\\d*)\\s+(" + RPT_UNITS + ")\\b");
  const vendorish = t => /^[A-Z]{2,}$/.test(t) && !/\d/.test(t) && !/^(ASSY|ASSEMBLY|MODEL|UNIT|KIT|SET|WHA|NBP|EDU|LG|THK|HD|OD|ID|SS|CRS|RTV|CLEAR|GRAY|GREY|BLACK|WHITE|RED|BLUE|BLU|GRN|GREEN|YEL|BRN|VIO|ORG|WHT|BLK|STRIPE|WIDE|ROUND|LEAD|PARTS|TUBE|SHEET|SPRING|PLATE|RING|WAX|SILICA)$/.test(t);
  const flush = () => { if (cur) { cur.desc = cur.desc.replace(/\s+/g, " ").trim(); rows.push(cur); } cur = null; };
  for (const raw of lines) {
    if (!raw) continue;
    const l = raw;
    if (/^Part No:/i.test(l)) { flush(); expectPN = true; continue; }
    if (expectPN) {
      const tok = l.split(/\s+/)[0];
      if (isPN(tok)) { parent = tok; parentRev = "-"; sections.push(parent); expectPN = false; continue; }
    }
    const revM = l.match(/Revision:\s*([A-Z0-9-]+(?:\s*\(\d+\))?)/i);
    if (revM && parent) { parentRev = revM[1].split(/\s/)[0]; const pr = rows.find(r => r.pn === parent); if (pr) pr.rev = parentRev; continue; }
    if (SKIP_RE.test(l)) { if (/^For Assy|^Cautions?/i.test(l) && cur) cur.rem = (cur.rem + " " + l).trim(); else flush(); continue; }
    const m = l.match(rowRe);
    if (m && isPN(m[1]) && parent) {
      flush();
      let desc = m[2].trim(); let rem = "";
      const dt = desc.split(/\s+/);
      if (dt.length > 1 && vendorish(dt[dt.length - 1])) { rem = "Vendor: " + dt.pop(); desc = dt.join(" "); }
      cur = { parent, find: "", qty: numQty(m[3]), uom: m[4], pn: m[1], rev: "-", desc, mb: /^Assy\s/.test(l) ? "" : "Buy", mat: "", rem };
      continue;
    }
    if (cur && !/^\$|^[\d,\s.$-]+$/.test(l) && !/^(Assy)$/.test(l)) cur.desc += " " + l;
  }
  flush();
  // parent parts that never appear as children: synthesize top rows
  const childPNs = new Set(rows.map(r => r.pn));
  const secRows = [];
  sections.forEach(s => { if (!childPNs.has(s)) secRows.push({ parent: "", find: "", qty: "1", uom: "EA", pn: s, rev: "-", desc: "", mb: "Make", mat: "", rem: "" }); });
  return { mode: "parent", rows: [...secRows, ...rows], meta: { format: "standard", sections, notes } };
}

/* ---- C. Report grids (xls) — column-true parsing ---- */
function parseReportGrid(grid) {
  const compact = row => (row || []).map(c => String(c == null ? "" : c).trim());
  const nonEmpty = row => compact(row).filter(Boolean);
  const flat = grid.map(nonEmpty);
  const all = flat.flat().join(" ");
  const isLevelCell = c => /^\d\*{2,}$/.test(c) || /^\d\*{2,}\s/.test(c);
  const hasLevels = flat.some(r => r[0] && isLevelCell(r[0]));
  const isStandard = /Bill of Materials/i.test(all) && /Part No:/i.test(all);
  if (!hasLevels && !isStandard) return null;
  const rows = []; const notes = []; const sections = [];
  const unitRe = new RegExp("^(" + RPT_UNITS + ")$");
  if (hasLevels) {
    for (const r of flat) {
      if (!r.length) continue;
      const lv = r[0].match(/^(\d)\*{2,}/);
      if (r[0] === "NA") { notes.push("Tooling (not BOM): " + (r[1] || "") + " " + (r.slice(2).find(c => /[a-z]/i.test(c) && !unitRe.test(c)) || "")); continue; }
      if (!lv) continue;
      const cells = r[0].includes(" ") ? [r[0].replace(/^\d\*+\s*/, ""), ...r.slice(1)] : r.slice(1);
      if (!cells.length || !isPN(cells[0])) continue;
      const pn = cells[0];
      let rev = "-", di = 1;
      if (cells[1] && /^([A-Z]|-|[A-Z]\d?)$/.test(cells[1]) && cells[1].length <= 2) { rev = cells[1]; di = 2; }
      let desc = "", qty = "1", uom = "EA";
      for (let i = di; i < cells.length; i++) {
        const c = cells[i];
        if (unitRe.test(c)) { uom = c; break; }
        if (/^[\d,]+\.?\d*$/.test(c) && desc) { qty = numQty(c); continue; }
        if (/[A-Za-z]/.test(c) && !/^YES$|^NO$/.test(c)) desc = desc ? desc + " " + c : c;
      }
      rows.push({ level: +lv[1], parent: null, find: "", qty, uom, pn, rev, desc, mb: "", mat: "", rem: "" });
    }
    return { mode: "level", rows, meta: { format: "kitting-grid", notes } };
  }
  // standard grid: track sections; child rows have qty+unit adjacency
  let parent = null;
  for (let ri = 0; ri < grid.length; ri++) {
    const rc = compact(grid[ri]);
    const ne = rc.filter(Boolean);
    if (!ne.length) continue;
    const pi = rc.findIndex(c => /^Part No:$/i.test(c));
    if (pi >= 0) { const v = rc.slice(pi + 1).find(c => isPN(c)); if (v) { parent = v; sections.push(parent); } continue; }
    const rvi = rc.findIndex(c => /^Revision:$/i.test(c));
    if (rvi >= 0 && parent) { const v = rc.slice(rvi + 1).find(Boolean); if (v) { const pr = rows.find(x => x.pn === parent); if (pr) pr.rev = v.split(/\s/)[0]; } }
    if (!parent) continue;
    const GRID_SKIP = /^(Page \d+ of|Report Generated|Island Components|Bill of Materials|TOTAL MATERIAL|Type$|Number$|Description$|Vendor$|Quantity$|Unit$|Weight$|Cost$|Amount$)/i;
    if (ne.some(c => GRID_SKIP.test(c))) continue;
    if (ne.some(c => /^For Assy/i.test(c))) continue;
    // find qty+unit pair
    let qi = -1;
    for (let i = 0; i < ne.length - 1; i++) if (/^[\d,]+\.?\d*$/.test(ne[i]) && unitRe.test(ne[i + 1])) { qi = i; break; }
    if (qi < 0) continue;
    const isAssy = ne[0] === "Assy";
    const body = ne.slice(isAssy ? 1 : 0, qi);
    if (!body.length || !isPN(body[0])) continue;
    const pn = body[0];
    let rev = "-", bi = 1;
    if (body[1] && body[1].length <= 8 && /^[A-Z](\s*\(\d+\))?$|^-$/.test(body[1])) { rev = body[1].split(/\s/)[0]; bi = 2; }
    let desc = "", rem = "";
    for (let i = bi; i < body.length; i++) {
      const c = body[i];
      if (/^[A-Z]{2,}$/.test(c) && !/\d/.test(c) && desc && c.length <= 12 && i === body.length - 1 && !/^(ASSY|ASSEMBLY|MODEL|UNIT|KIT|SET|WHA|NBP|EDU|LG|THK|HD|OD|ID|SS|CRS|RTV|CLEAR|GRAY|GREY|BLACK|WHITE|RED|BLUE|BLU|GRN|GREEN|YEL|BRN|VIO|ORG|WHT|BLK|STRIPE|WIDE|ROUND|LEAD|PARTS|TUBE|SHEET|SPRING|PLATE|RING|WAX|SILICA)$/.test(c)) { rem = "Vendor: " + c; continue; }
      desc = desc ? desc + " " + c : c;
    }
    rows.push({ parent, find: "", qty: numQty(ne[qi]), uom: ne[qi + 1], pn, rev, desc, mb: isAssy ? "" : "Buy", mat: "", rem });
  }
  const childPNs = new Set(rows.map(r => r.pn));
  const secRows = [];
  sections.forEach(s => { if (!childPNs.has(s)) secRows.push({ parent: "", find: "", qty: "1", uom: "EA", pn: s, rev: "-", desc: "", mb: "Make", mat: "", rem: "" }); });
  return { mode: "parent", rows: [...secRows, ...rows], meta: { format: "standard-grid", sections, notes } };
}

function reportToGrid(parsed) {
  const hdrKey = parsed.mode === "level" ? "Level" : "Parent";
  let lvl0 = 0;
  if (parsed.mode === "level" && parsed.rows.length) lvl0 = Math.min(...parsed.rows.map(r => +r.level || 0));
  const headers = [hdrKey, "Find", "Qty", "UOM", "PartNumber", "Rev", "Description", "MakeBuy", "Material", "Remarks"];
  const rows = parsed.rows.map(r => [
    parsed.mode === "level" ? String((+r.level || 0) - lvl0) : (r.parent || ""),
    r.find || "", r.qty || "1", r.uom || "EA", r.pn, r.rev || "-", r.desc || "", r.mb || "", r.mat || "", r.rem || ""]);
  return [headers, ...rows];
}

/* =====================================================================
   DRAWING MODULE — deterministic extraction from drawing text
   (vector-PDF text or OCR output). Finds title-block PN + parts-list rows.
   ===================================================================== */
const DRAW_PN = "[A-Z]{2,5}-?\\d{3,6}[A-Z0-9-]*";
function extractFromDrawingText(text) {
  const lines = String(text || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let parent = null, title = null;
  for (const l of lines) {
    const m = l.match(new RegExp("(?:DWG|DRAWING|PART)\\s*(?:NO|NUMBER|#)?\\s*[.:]?\\s*(" + DRAW_PN + ")", "i"));
    if (m) { parent = m[1].toUpperCase(); break; }
  }
  for (const l of lines) {
    const m = l.match(/(?:TITLE|NAME)\s*[.:]\s*(.{4,60})/i);
    if (m) { title = m[1].trim(); break; }
  }
  const rows = []; const rejects = [];
  for (const l of lines) {
    // patterns: "ITEM QTY PN DESCRIPTION" | "FIND PN DESC QTY [UOM]" | "PN DESC QTY"
    let m = l.match(new RegExp("^(\\d{1,3})[\\s.]+(\\d{1,3}|AR)\\s+(" + DRAW_PN + ")\\s+(.+)$", "i"));
    if (m) { rows.push({ find: m[1], qty: m[2].toUpperCase(), uom: "EA", pn: m[3].toUpperCase(), rev: "-", desc: m[4].replace(/\s{2,}.*$/, "").trim(), mb: "", mat: "", rem: "" }); continue; }
    m = l.match(new RegExp("^(\\d{1,3})[\\s.]+(" + DRAW_PN + ")\\s+(.+?)\\s+(\\d{1,3}|AR)\\s*(EA|GM|LB|IN|FT|PC|SET)?\\s*$", "i"));
    if (m) { rows.push({ find: m[1], qty: m[4].toUpperCase(), uom: (m[5] || "EA").toUpperCase(), pn: m[2].toUpperCase(), rev: "-", desc: m[3].trim(), mb: "", mat: "", rem: "" }); continue; }
    m = l.match(new RegExp("^(" + DRAW_PN + ")\\s+(.+?)\\s+(\\d{1,3}|AR)\\s*(EA|GM|LB)?\\s*$", "i"));
    if (m && m[1].toUpperCase() !== parent) { rows.push({ find: String((rows.length + 1) * 10), qty: m[3].toUpperCase(), uom: (m[4] || "EA").toUpperCase(), pn: m[1].toUpperCase(), rev: "-", desc: m[2].trim(), mb: "", mat: "", rem: "" }); continue; }
    if (new RegExp(DRAW_PN).test(l) && !/^(?:DWG|DRAWING|PART|TITLE|SHEET|SCALE|REV)/i.test(l)) rejects.push(l);
  }
  // PN sanity: flag OCR-suspect characters
  const suspect = rows.filter(r => /[OIl]/.test(r.pn.replace(/[A-Z]{2,5}-/, ""))).map(r => r.pn);
  return { parent, title, rows, rejects, suspect };
}

/* =====================================================================
   ADAPTER HOOKS — Tesseract OCR + local LLM (Ollama-style endpoint).
   Generation path never uses these; import-edge assist only.
   Desktop build / power users register: window.docworksAdapters = {  (legacy alias: motoflowAdapters)
     tesseract: { recognize: async (imageBlobOrUrl) => "text" },
     pdfText:   { extract:   async (arrayBuffer)    => "text" },
   }
   Local LLM: any OpenAI/Ollama-compatible endpoint reachable from the page.
   ===================================================================== */
function getAdapter(name) {
  try { if (typeof window === "undefined") return null;
    return (window.docworksAdapters && window.docworksAdapters[name]) || (window.motoflowAdapters && window.motoflowAdapters[name]) || null; }
  catch (e) { return null; }
}
async function ocrImage(fileOrBlob) {
  const ad = getAdapter("tesseract");
  if (ad && ad.recognize) return await ad.recognize(fileOrBlob);
  const T = await ensureTesseract(); // lazy-load from CDN in the browser prototype
  const res = await T.recognize(fileOrBlob, "eng");
  return res && res.data ? res.data.text : "";
}
/* lazy script loader (cdnjs) — used only in the browser prototype; the desktop
   build bundles pdf.js + tesseract.js locally so nothing leaves the machine. */
function loadScript(url) {
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = url; s.onload = () => res(); s.onerror = () => rej(new Error("Failed to load " + url));
    document.head.appendChild(s);
  });
}
async function ensurePdfJs() {
  if (typeof window === "undefined") throw new Error("No browser environment.");
  if (window.pdfjsLib) return window.pdfjsLib;
  const base = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/";
  await loadScript(base + "pdf.min.js");
  if (!window.pdfjsLib) throw new Error("pdf.js did not initialize.");
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = base + "pdf.worker.min.js";
  return window.pdfjsLib;
}
async function ensureTesseract() {
  if (typeof window === "undefined") throw new Error("No browser environment.");
  if (window.Tesseract) return window.Tesseract;
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/4.1.1/tesseract.min.js");
  if (!window.Tesseract) throw new Error("tesseract.js did not initialize.");
  return window.Tesseract;
}
/* reconstruct reading-order lines from pdf.js text items (group by y, sort by x) */
function pdfItemsToLines(items) {
  const rows = new Map();
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    const y = Math.round(it.transform[5] / 3) * 3; // 3pt y-tolerance
    if (!rows.has(y)) rows.set(y, []);
    rows.get(y).push({ x: it.transform[4], s: it.str });
  }
  return [...rows.entries()].sort((a, b) => b[0] - a[0])
    .map(([, cells]) => cells.sort((a, b) => a.x - b.x).map(c => c.s).join("  ").replace(/\s{2,}/g, "  ").trim());
}
async function pdfToText(arrayBuffer, onProgress) {
  const ad = getAdapter("pdfText");
  if (ad && ad.extract) return await ad.extract(arrayBuffer);
  const pdfjs = await ensurePdfJs();
  const doc = await pdfjs.getDocument({ data: arrayBuffer }).promise;
  let all = [];
  for (let i = 1; i <= doc.numPages; i++) {
    if (onProgress) onProgress(`reading page ${i}/${doc.numPages}`);
    const page = await doc.getPage(i);
    const tc = await page.getTextContent();
    all = all.concat(pdfItemsToLines(tc.items), [""]);
  }
  const text = all.join("\n");
  if (text.replace(/\s/g, "").length >= 30) return text;
  // no text layer -> scanned drawing: rasterize pages and OCR them
  if (onProgress) onProgress("no text layer — running OCR on scanned pages");
  const T = await ensureTesseract();
  let ocrAll = [];
  for (let i = 1; i <= Math.min(doc.numPages, 6); i++) {
    if (onProgress) onProgress(`OCR page ${i}/${Math.min(doc.numPages, 6)}`);
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 2.2 });
    const cv = document.createElement("canvas");
    cv.width = vp.width; cv.height = vp.height;
    await page.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
    const res = await T.recognize(cv, "eng");
    ocrAll.push((res && res.data && res.data.text) || "");
  }
  return ocrAll.join("\n");
}
async function llmNormalize(rawText, cfg) {
  // cfg: {url, model}. Ollama /api/chat compatible. Returns CSV text or throws.
  const body = {
    model: cfg.model, stream: false,
    messages: [
      { role: "system", content: "You convert messy engineering drawing parts-list text into clean CSV with header: Find,Qty,UOM,PartNumber,Rev,Description. Output ONLY the CSV, no commentary. Preserve part numbers exactly; do not invent rows." },
      { role: "user", content: rawText.slice(0, 8000) },
    ],
  };
  const r = await fetch(cfg.url.replace(/\/$/, "") + "/api/chat", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  if (!r.ok) throw new Error("LLM endpoint HTTP " + r.status);
  const j = await r.json();
  const out = (j.message && j.message.content) || (j.choices && j.choices[0] && j.choices[0].message.content) || "";
  if (!/PartNumber|,/.test(out)) throw new Error("LLM returned no CSV.");
  return out.replace(/```[a-z]*\n?|```/g, "").trim();
}

/* =====================================================================
   UI THEME
   ===================================================================== */
const C = {
  navy: "#1F3864", navy2: "#2E5395", paper: "#FAFAF8", ink: "#1A1A1E",
  line: "#D8D8D4", hold: "#FFF2CC", holdInk: "#8A6D00", stamp: "#2E6B3E",
  warn: "#FBE4D5", warnInk: "#B03A00", note: "#E2EFD9", backdrop: "#E9E9E4",
  ltblue: "#D9E2F1", gray: "#F2F2F2",
};
const MONO = '"Consolas","JetBrains Mono",ui-monospace,monospace';

/* =====================================================================
   OUTPUT PROFILES — EZ (default) and Island Components Group.
   Profile controls branding, doc-number scheme, column labels, op-verb
   vocabulary, and which document set is emitted. Generation logic is
   shared; profiles are presentation + naming only.
   ===================================================================== */
const ISLAND_LOGO = "data:image/png;base64,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";
const PROFILES = {
  ez: {
    id: "ez", label: "EZ Motors (default)",
    company: "EZ MOTORS, INC.", companyShort: "EZ MOTORS",
    address: "AEROSPACE ELECTROMECHANICAL SYSTEMS", cage: "CAGE CODE: 0Z9K8",
    docNo: { F: t => "F-" + t, P: t => "P-" + t, TRV: pn => "TRV-" + pn, WI: pn => "WI-" + pn },
    wiKind: "WORK INSTRUCTION", wiRef: pn => "WI-" + pn,
    qaField: false, espMode: false,
    travelerCols: ["Op", "Dept", "Operation / Detailed Work Instruction", "Acceptance / Data to Record", "Qty Pass", "Qty Rej", "Operator / Date", "QA / Date"],
    mapVerb: null,
  },
  island: {
    id: "island", label: "Island Components Group",
    company: "ISLAND COMPONENTS GROUP, INC.", companyShort: "ISLAND COMPONENTS",
    address: "210 Marcus Blvd. · Hauppauge, N.Y. 11788", cage: "Tel (631) 563-4224 · Fax (631) 563-4363",
    // Island uses ESP procedure numbers as the WI identity; travelers are Job Travelers.
    docNo: { F: t => t + " FAMILY TREE", P: t => "PARTS LIST " + t, TRV: pn => "Job Traveler " + pn, WI: (pn, esp) => (esp || "ESP-*") },
    wiKind: "PROCESS PROCEDURE", wiRef: (pn, esp) => (esp || "ESP-*"),
    qaField: true, espMode: true,
    travelerCols: ["Step No", "Dept", "Work Cntr", "Operation Description", "Qty Accepted", "Qty Reject", "Complete By / Date", "Comments"],
    // map generic op titles -> Island work-center action verbs (departments preserved separately)
    mapVerb: title => {
      const T = title.toLowerCase();
      if (/kit/.test(T)) return "KIT";
      if (/inspect|inspection|verify|final qa/.test(T)) return "INSPECT";
      if (/wind/.test(T)) return "WIND";
      if (/insert coil|insert/.test(T)) return "INSERT";
      if (/rotation|test|hi-?pot|resistance|electrical|impregnation test/.test(T)) return "TEST";
      if (/connect|lead|solder|terminat/.test(T)) return "CONNECT";
      if (/tape|teflon/.test(T)) return "TAPE";
      if (/lace|lacing/.test(T)) return "INSERT";
      if (/form/.test(T)) return "FORM";
      if (/varnish|impregnat/.test(T)) return "IMPREG";
      if (/clean|visual/.test(T)) return "CLEAN";
      if (/mark/.test(T)) return "MARK";
      if (/bond|stack|assemble|install/.test(T)) return "ASSEMBLE";
      if (/grind|machine/.test(T)) return "MACHINE";
      if (/stock|preserve|package|ship/.test(T)) return "MOVE";
      if (/cure/.test(T)) return "CURE";
      return "MFG";
    },
  },
};
function activeProfile(id) { return PROFILES[id] || PROFILES.ez; }
// Island departments derived from our dept text (keep real departments per user)
function islandDept(dept) {
  const d = (dept || "").toUpperCase();
  if (/QA|QUALITY/.test(d)) return "QA";
  if (/INSPECT|INSP/.test(d)) return "INSP";
  if (/TEST/.test(d)) return "TEST";
  if (/STORES|STOCK|KIT/.test(d)) return "MFG";
  if (/MACHINE|GRIND|PROC|WIND|ASSEMBLY|MFG/.test(d)) return "MFG";
  return "MFG";
}
/* customer auto-detect from BOM descriptions/remarks (e.g. "Mitutoyo") */
const KNOWN_CUSTOMERS = ["Mitutoyo", "Renishaw", "Alpine", "Genesis", "Brewster", "Tower", "Wiremasters", "Century Spring"];
function detectCustomer(bom) {
  const hay = Object.values(bom.parts).map(p => (p.desc || "") + " " + (p.rem || "")).join(" ");
  for (const c of KNOWN_CUSTOMERS) if (new RegExp(c, "i").test(hay)) return { value: c, guessed: true };
  const m = hay.match(/\bCUST(?:OMER)?[:\s]+([A-Z][A-Za-z]{3,})/);
  if (m) return { value: m[1], guessed: true };
  return { value: "", guessed: false };
}
/* ESP number resolution: per-assembly override map -> default ESP-* */
function espFor(pn, espByPn) { return (espByPn && espByPn[pn]) ? espByPn[pn] : "ESP-*"; }
/* yellow highlight wrapper for auto-populated / uncertain fields (Island) */
const HL = { background: "#FFF200", padding: "0 3px", fontWeight: 700 };
function Y({ children, on }) { return on ? <span style={HL}>{children}</span> : <>{children}</>; }
/* pull fixture / tool callouts and material PNs referenced in an op set (for ESP auto-populate) */
function extractToolsAndDocs(ops) {
  const tools = new Set(), docs = new Set(), specs = new Set();
  const all = ops.map(o => [o.title, o.text, (o.sub || []).join(" "), o.record, o.accept].join(" ")).join(" ");
  (all.match(/\b(?:FF|TL|FIX)-\d{2,4}[A-Z0-9-]*/g) || []).forEach(t => tools.add(t));
  (all.match(/\bESP-\d{2,4}[A-Z0-9-]*/g) || []).forEach(d => docs.add(d));
  (all.match(/\bdrawing\s+([A-Z]{1,4}-?\d{2,5}[A-Z0-9-]*)/gi) || []).forEach(d => docs.add(d.replace(/drawing\s+/i, "")));
  return { tools: [...tools], docs: [...docs], specs: [...specs] };
}

function DocHeader({ title, docNo, pn, extra, m, company }) {
  const cell = { padding: "5px 8px", borderRight: `1px solid ${C.navy}`, fontSize: 11 };
  const grid = { display: "grid", gridTemplateColumns: "120px 1fr 105px 45px 115px", border: `1.5px solid ${C.navy}` };
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={grid}>
        <div style={{ ...cell, background: C.navy, color: "#fff", fontWeight: 800, letterSpacing: ".06em", fontSize: 12 }}>{company || "EZ MOTORS"}</div>
        <div style={{ ...cell, background: C.navy, color: "#fff", fontWeight: 700 }}>{title}</div>
        <div style={{ ...cell, background: C.ltblue, fontWeight: 700, fontSize: 9, textTransform: "uppercase" }}>Doc No.</div>
        <div style={{ ...cell, background: C.ltblue, fontWeight: 700, fontSize: 9, textTransform: "uppercase" }}>Rev</div>
        <div style={{ ...cell, background: C.ltblue, fontWeight: 700, fontSize: 9, textTransform: "uppercase", borderRight: "none" }}>Status</div>
      </div>
      <div style={{ ...grid, borderTop: "none" }}>
        <div style={cell} />
        <div style={{ ...cell, fontFamily: MONO }}>{pn}{extra ? ` — ${extra}` : ""}</div>
        <div style={{ ...cell, fontFamily: MONO }}>{docNo}</div>
        <div style={{ ...cell, fontFamily: MONO }}>{m.rev || "1"}</div>
        <div style={{ ...cell, fontSize: 8.5, borderRight: "none" }}>SAMPLE / UNRELEASED</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", border: `1px solid ${C.line}`, borderTop: "none" }}>
        {["Work Order", "Serial / Lot", "Program", "Issue Date", "Prepared", "Approved"].map((k, i) => (
          <div key={k} style={{ padding: "3px 8px", borderRight: i < 5 ? `1px solid ${C.line}` : "none", background: C.gray, fontWeight: 700, fontSize: 9, textTransform: "uppercase", color: "#555" }}>{k}</div>
        ))}
        {[m.wo, m.sn, m.prog, m.date, "Engineering", "Quality"].map((v, i) => (
          <div key={i} style={{ padding: "3px 8px", borderRight: i < 5 ? `1px solid ${C.line}` : "none", fontFamily: i < 2 || i === 3 ? MONO : "inherit", fontSize: 10.5 }}>{v}</div>
        ))}
      </div>
    </div>
  );
}
function Sheet({ children, wide }) {
  return <div style={{ background: "#fff", width: "100%", maxWidth: wide ? 1360 : 850, boxShadow: "0 2px 14px rgba(0,0,0,.13)", padding: wide ? "24px 28px" : "30px 34px", fontSize: 12.5, lineHeight: 1.45 }}>{children}</div>;
}
function H3({ children }) { return <h3 style={{ fontSize: 13, color: C.navy, letterSpacing: ".04em", margin: "16px 0 8px", fontWeight: 700 }}>{children}</h3>; }
function Intro({ children }) { return <p style={{ fontSize: 11.5, color: "#666", fontStyle: "italic", marginBottom: 10 }}>{children}</p>; }
function Callout({ k, v }) {
  const map = { WARNING: [C.warnInk, C.warn], CAUTION: [C.holdInk, C.hold], NOTE: [C.stamp, C.note] };
  const [ink, bg] = map[k] || map.NOTE;
  return (
    <div style={{ display: "grid", gridTemplateColumns: "85px 1fr", border: `1px solid ${C.line}`, margin: "8px 0", fontSize: 11 }}>
      <div style={{ padding: 6, fontWeight: 800, textAlign: "center", background: bg, color: ink, display: "flex", alignItems: "center", justifyContent: "center" }}>{k}</div>
      <div style={{ padding: "6px 9px", background: bg }}>{v}</div>
    </div>
  );
}
function Foot({ company }) { return <div style={{ textAlign: "center", color: "#999", fontSize: 9.5, marginTop: 18, letterSpacing: ".04em" }}>UNCONTROLLED WHEN PRINTED | {company || "EZ MOTORS"} PROPRIETARY | AS9100 CONTROLLED DOCUMENT (SAMPLE)</div>; }

/* ---- Family tree doc — drawing-sheet style (layout v2: ladders, min scale, sheet sets) ---- */
function isAssemblyLike(part) { return /assembl|assy|\bkit\b/i.test(part.desc || ""); }
/* layout v2 — readable letter-size family trees
   - Leaf ladder: >=4 leaf children stack vertically (1-2 columns) instead of one wide row
   - MAXW per sheet enforces minimum print scale (PN >= ~8pt on letter landscape)
   - Pagination: if a tree exceeds MAXW, top sheet collapses deep assemblies to
     "SEE SHEET n" refs and each gets its own sheet
*/
const NW = 118, NH = 60, GX = 12, ROWH = 116, STACK_PITCH = NH + 12;
const LEAF_STACK_MIN = 4;   // stack when this many leaf children
const MAXW = 980;           // px; ~1:1 scale on letter landscape -> PN ~8pt
const MAXH_MAIN = 440;      // px; drawing height budget, main sheet (title+legend share the page)
const MAXH_SUB = 560;       // px; sub-sheets have more room
// SHEET SIZES: "letter" = readable multi-sheet set; "tabloid" = one 11x17 sheet, compact
const SHEET_SIZES = {
  letter:  { NW: 118, NH: 60, GX: 12, ROWH: 116, descLines: 3, descChars: 19, fontPN: 11, MAXW: 980, usableW: 984, oneSheet: false },
  tabloid: { NW: 120, NH: 52, GX: 30, ROWH: 96,  descLines: 1, descChars: 22, fontPN: 11, MAXW: 100000, usableW: 1540, oneSheet: true, bulletLeaves: true, bulletW: 250, bulletLH: 13, VGAP: 34 },
};
function sheetDims(sizeId) { return SHEET_SIZES[sizeId] || SHEET_SIZES.letter; }

/* BULLET LAYOUT (ICG house style): assemblies are boxed; their component (leaf)
   children are listed as an indented bullet list directly beneath the box with qty.
   Sub-assemblies branch as sibling boxes. Vertical placement is dynamic (a node's
   children start below the node's box + its bullet list). Used for the 11x17 sheet. */
function layoutBulletTree(bom, excluded, top, D) {
  // Components sit in a RIGHT-HAND column beside the assembly box (not below it),
  // so the sub-assembly connector can exit the box bottom cleanly without crossing
  // the component list. Each node occupies a block = [box | components column].
  const NW = D.NW, NH = D.NH, GX = D.GX, BW = D.bulletW || 250, BLH = D.bulletLH || 13, VGAP = D.VGAP || 34;
  const COLGAP = 10;        // gap between box and its component column
  const BOXGAP = 34;        // gap between sibling assembly blocks
  function build(pn, row, depth, seen) {
    const cyc = seen.has(pn);
    const kidRows = cyc ? [] : (bom.children[pn] || []).filter(k => !excluded[k.pn]);
    const ns = new Set(seen); ns.add(pn);
    const kids = kidRows.map(k => build(k.pn, k, depth + 1, ns));
    const asmK = kids.filter(k => k.kids.length);
    const leafK = kids.filter(k => !k.kids.length);
    return { pn, row, depth, part: bom.parts[pn] || { pn, desc: "", rev: "-", mb: "" }, kids, asmK, leafK };
  }
  const root = build(top, null, 0, new Set());
  // blockW = box + (component column to the right, if any)
  function width(n) {
    n.hasBullets = n.leafK.length > 0;
    n.compW = n.hasBullets ? BW : 0;
    n.blockW = NW + (n.hasBullets ? COLGAP + n.compW : 0);
    n.asmK.forEach(width);
    const childrenW = n.asmK.reduce((s, k) => s + k.blockW, 0) + BOXGAP * Math.max(0, n.asmK.length - 1);
    n.w = Math.max(n.blockW, childrenW);
    return n.w;
  }
  width(root);
  // block height = max(box height, component column height)
  const compH = n => n.hasBullets ? (n.leafK.length * BLH + 8) : 0;
  function placeX(n, x) {
    // n.boxX = left edge of the assembly box; box sits at the left of the block
    if (!n.asmK.length) { n.boxX = x + (n.w - n.blockW) / 2; n.x = n.boxX + NW / 2; return; }
    const childrenW = n.asmK.reduce((s, k) => s + k.blockW, 0) + BOXGAP * Math.max(0, n.asmK.length - 1);
    let cx = x + Math.max(0, (n.w - childrenW) / 2);
    n.asmK.forEach(k => { placeX(k, cx); cx += k.blockW + BOXGAP; });
    // parent box centers over the span of child BOXES (not their component columns)
    const boxCenters = n.asmK.map(k => k.x);
    const mid = (Math.min(...boxCenters) + Math.max(...boxCenters)) / 2;
    n.boxX = mid - NW / 2;
    n.x = mid;
  }
  placeX(root, 0);
  let maxY = 0;
  (function placeY(n, topY) {
    n.y = topY;
    n.blockH = Math.max(NH, compH(n));
    const below = topY + n.blockH;
    maxY = Math.max(maxY, below);
    if (n.asmK.length) { const ct = below + VGAP; n.asmK.forEach(k => placeY(k, ct)); }
  })(root, 0);
  const flat = []; (function w(n) { flat.push(n); n.asmK.forEach(w); })(root);
  return { root, flat, totalW: root.w, totalH: maxY, bullet: true, rightColumn: true };
}

function layoutTree2(bom, excluded, top, collapse, stackMode, D) {
  D = D || sheetDims("letter");
  const NW = D.NW, NH = D.NH, GX = D.GX, ROWH = D.ROWH, STACK_PITCH = D.NH + 12;
  // collapse: Set of pns to render as collapsed boxes (no children drawn)
  // stackMode: only stack leaf ladders when true (width-driven, decided by planner)
  collapse = collapse || new Set();
  function build(pn, row, depth) {
    const collapsed = collapse.has(pn) && depth > 0;
    const kidRows = collapsed ? [] : (bom.children[pn] || []).filter(k => !excluded[k.pn]);
    const kids = kidRows.map(k => build(k.pn, k, depth + 1));
    const asmK = kids.filter(k => k.kids.length || k.collapsed);
    const leafK = kids.filter(k => !k.kids.length && !k.collapsed);
    return { pn, row, depth, part: bom.parts[pn] || { pn, desc: "", rev: "-", mb: "" }, kids, asmK, leafK, collapsed };
  }
  const root = build(top, null, 0);

  function width(n) {
    if (!n.kids.length) { n.w = NW; n.inline = []; n.stacked = []; return n.w; }
    const doStack = stackMode && (n.leafK.length >= LEAF_STACK_MIN || (n.asmK.length > 0 && n.leafK.length >= 3));
    n.stacked = doStack ? n.leafK : [];
    n.inline = doStack ? n.asmK : n.kids;
    n.inline.forEach(width);
    n.stacked.forEach(k => { k.w = NW; k.inline = []; k.stacked = []; });
    let w = n.inline.reduce((s, k) => s + k.w, 0) + GX * Math.max(0, n.inline.length - 1);
    if (n.stacked.length) {
      n.stackCols = n.stacked.length > 8 ? 3 : n.stacked.length > 4 ? 2 : 1;
      n.stackRows = Math.ceil(n.stacked.length / n.stackCols);
      n.stackW = n.stackCols * NW + (n.stackCols - 1) * GX + 14; // +rail allowance
      w += (n.inline.length ? GX : 0) + n.stackW;
    }
    n.w = Math.max(NW, w);
    return n.w;
  }
  width(root);

  function place(n, x) {
    if (!n.kids.length || (!n.inline.length && !n.stacked.length)) { n.x = x + n.w / 2; return; }
    const contentW = n.inline.reduce((s, k) => s + k.w, 0) + GX * Math.max(0, n.inline.length - 1)
      + (n.stacked.length ? (n.inline.length ? GX : 0) + n.stackW : 0);
    let cx = x + Math.max(0, (n.w - contentW) / 2);
    n.inline.forEach(k => { place(k, cx); cx += k.w + GX; });
    if (n.stacked.length) {
      n.stackX = cx + 14; // rail sits in the 14px allowance
      n.stacked.forEach((k, i) => {
        const r = Math.floor(i / n.stackCols), c = i % n.stackCols;
        k.x = n.stackX + c * (NW + GX) + NW / 2;
        k.stackRow = r; k.isStacked = true;
      });
    }
    const centers = [
      ...n.inline.map(k => k.x),
      ...(n.stacked.length ? [n.stackX + (n.stackCols * NW + (n.stackCols - 1) * GX) / 2] : []),
    ];
    n.x = centers.length ? (Math.min(...centers) + Math.max(...centers)) / 2 : x + n.w / 2;
  }
  place(root, 0);

  const flat = []; let maxY = 0;
  (function walk(n) {
    n.y = n.depth * ROWH + (n.isStacked ? n.stackRow * STACK_PITCH : 0);
    maxY = Math.max(maxY, n.y + NH);
    flat.push(n);
    n.inline.forEach(walk);
    n.stacked.forEach(walk);
  })(root);
  return { root, flat, totalW: root.w, totalH: maxY };
}

/* sheet planner: plain layout -> stacked layout -> paginate; returns [{top, collapse, refs, layout, sheetNo}] */
function planSheets(bom, excluded, top, maxH, D) {
  D = D || sheetDims("letter");
  // ONE-SHEET (11x17) mode: whole tree on a single sheet, stacked ladders on, no pagination
  if (D.oneSheet) {
    let L = D.bulletLeaves ? layoutBulletTree(bom, excluded, top, D) : layoutTree2(bom, excluded, top, null, true, D);
    return [{ top, collapse: new Set(), refs: {}, layout: L, sheetNo: 1, oneSheet: true }];
  }
  maxH = maxH || MAXH_MAIN;
  const fits = L => L.totalW <= D.MAXW && L.totalH <= maxH;
  const plain = layoutTree2(bom, excluded, top, null, false, D);
  if (fits(plain)) return [{ top, collapse: new Set(), refs: {}, layout: plain, sheetNo: 1 }];
  const stacked = layoutTree2(bom, excluded, top, null, true, D);
  if (fits(stacked)) return [{ top, collapse: new Set(), refs: {}, layout: stacked, sheetNo: 1 }];
  const lvl1Asms = (bom.children[top] || []).filter(k => !excluded[k.pn] && (bom.children[k.pn] || []).some(c => !excluded[c.pn])).map(k => k.pn);
  const collapse = new Set(lvl1Asms);
  let topLayout = layoutTree2(bom, excluded, top, collapse, false, D);
  if (topLayout.totalW > D.MAXW || topLayout.totalH > maxH) topLayout = layoutTree2(bom, excluded, top, collapse, true, D);
  const sheets = [{ top, collapse, refs: {}, layout: topLayout }];
  lvl1Asms.forEach(pn => sheets.push(...planSheets(bom, excluded, pn, MAXH_SUB, D)));
  sheets.forEach((s, i) => s.sheetNo = i + 1);
  const byTop = {}; sheets.forEach(s => { if (byTop[s.top] === undefined) byTop[s.top] = s.sheetNo; });
  sheets.forEach(s => { s.refs = {}; s.collapse.forEach(pn => { if (byTop[pn]) s.refs[pn] = byTop[pn]; }); });
  return sheets;
}


function wrapText(s, maxChars, maxLines) {
  maxLines = maxLines || 3;
  const words = String(s || "").toUpperCase().split(/\s+/).filter(Boolean);
  const lines = []; let cur = "";
  for (const w of words) {
    if ((cur + " " + w).trim().length <= maxChars) cur = (cur + " " + w).trim();
    else { if (cur) lines.push(cur); cur = w.length > maxChars ? w.slice(0, maxChars - 1) + "…" : w; }
    if (lines.length === maxLines) break;
  }
  if (cur && lines.length < maxLines) lines.push(cur);
  return lines.slice(0, maxLines);
}

function SheetDrawing({ bom, sheet, purchased, fit, qaField }) {
  const L = sheet.layout;
  const D = sheet.D || sheetDims("letter");
  const NW = D.NW, NH = D.NH, GX = D.GX, ROWH = D.ROWH;
  const compact = D.oneSheet;
  const fPN = D.fontPN, fD = compact ? 7.5 : 7.2, fQ = compact ? 8 : 8, dLines = D.descLines, dChars = D.descChars;
  const fBullet = D.bulletFont || 9.5;
  const PADX = 20, PADY = 16, CALLOUT_H = compact ? 30 : 46;
  const callouts = L.flat.filter(n => !n.isStacked && n.inline && n.inline.length && n.inline.every(k => !k.kids.length && !k.collapsed) && !(n.stacked && n.stacked.length) && n.depth >= 1);
  const H = L.totalH + (callouts.length ? CALLOUT_H : 0) + 26 + PADY * 2;
  const W = L.totalW + PADX * 2;

  // ---- BULLET LAYOUT RENDER (ICG house style) ----
  if (L.bullet) {
    const BLH = D.bulletLH || 13, BW = D.bulletW || 250, COLGAP = 10;
    const FONT = '"Arial","Helvetica",sans-serif'; // bolder, more legible than mono for the tree
    const bels = [];
    const truncate = (s, nchar) => { s = String(s || "").toUpperCase(); return s.length > nchar ? s.slice(0, nchar - 1) + "…" : s; };
    // component-label char budget derived from column width (so nothing clips off the side)
    const compChars = Math.max(24, Math.floor((BW - 14) / 5.4));
    for (const n of L.flat) {
      const bx0 = PADX + (n.boxX != null ? n.boxX : n.x - NW / 2), ny = PADY + n.y;
      const isTop = n.depth === 0;
      const missing = !n.kids.length && isAssemblyLike(n.part) && !purchased[n.pn];
      const isPurch = !n.kids.length && isAssemblyLike(n.part) && purchased[n.pn];
      const key = n.pn + "_" + n.depth + "_" + Math.round(n.x);
      // assembly box
      bels.push(<rect key={"b" + key} x={bx0} y={ny} width={NW} height={NH} rx={3}
        fill={missing ? "#FFF9E8" : isTop ? "#EEF3FB" : "#fff"} stroke={missing ? "#B8860B" : isTop ? C.navy : C.navy2}
        strokeWidth={isTop ? 2.2 : 1.5} strokeDasharray={missing ? "5 3" : "none"} />);
      bels.push(<text key={"p" + key} x={bx0 + NW / 2} y={ny + 15} textAnchor="middle" fontFamily={FONT} fontSize={fPN} fontWeight={800} fill={isTop ? C.navy : C.navy2}>{n.pn}</text>);
      wrapText(n.part.desc, dChars, 2).forEach((l, li) => bels.push(<text key={"d" + key + li} x={bx0 + NW / 2} y={ny + 26 + li * 9} textAnchor="middle" fontFamily={FONT} fontSize={fD} fontWeight={600} fill="#333">{l}</text>));
      bels.push(<text key={"q" + key} x={bx0 + NW / 2} y={ny + NH - 5} textAnchor="middle" fontFamily={FONT} fontSize={fQ} fontWeight={700} fill="#111">{"QTY: " + (n.row ? n.row.qty : "1")}</text>);
      if (missing) bels.push(<text key={"m" + key} x={bx0 + NW / 2} y={ny + NH + 9} textAnchor="middle" fontFamily={FONT} fontSize={7} fontWeight={700} fill="#B8860B">▲ NO BOM</text>);
      if (isPurch) bels.push(<text key={"u" + key} x={bx0 + NW / 2} y={ny + NH + 9} textAnchor="middle" fontFamily={FONT} fontSize={7} fontWeight={700} fill="#666">(PURCHASED)</text>);
      // component list to the RIGHT of the box (right-hand column)
      if (n.hasBullets) {
        const colX = bx0 + NW + COLGAP;
        const by = ny + 2;
        // connector: from box right-mid to the column
        const midY = ny + NH / 2;
        bels.push(<line key={"cc" + key} x1={bx0 + NW} y1={midY} x2={colX - 3} y2={midY} stroke="#999" strokeWidth={1} />);
        // vertical rail down the column
        bels.push(<line key={"rail" + key} x1={colX} y1={by} x2={colX} y2={by + n.leafK.length * BLH} stroke="#bbb" strokeWidth={.8} />);
        n.leafK.forEach((k, i) => {
          const ly = by + i * BLH + 9;
          bels.push(<line key={"bl" + key + i} x1={colX} y1={ly - 3} x2={colX + 5} y2={ly - 3} stroke="#bbb" strokeWidth={.8} />);
          const qty = k.row ? k.row.qty : "1";
          const label = k.pn + "  " + truncate(k.part.desc, compChars) + "  (" + qty + ")";
          const km = !k.kids.length && isAssemblyLike(k.part) && !purchased[k.pn];
          bels.push(<text key={"bt" + key + i} x={colX + 9} y={ly} fontFamily={FONT} fontSize={fBullet} fontWeight={500} fill={km ? "#B8860B" : "#1A1A1E"}>{label}{km ? " ▲" : ""}</text>);
        });
      }
      // connector to sub-assemblies: exits the BOX BOTTOM (clean, no list crossing)
      if (n.asmK && n.asmK.length) {
        const busY = ny + NH + (D.VGAP || 34) / 2;
        const px = bx0 + NW / 2;
        bels.push(<line key={"v" + key} x1={px} y1={ny + NH} x2={px} y2={busY} stroke={C.navy} strokeWidth={1.2} />);
        const ends = [px];
        n.asmK.forEach(k => { const kpx = PADX + (k.boxX != null ? k.boxX + NW / 2 : k.x); ends.push(kpx); bels.push(<line key={"c" + key + k.pn} x1={kpx} y1={busY} x2={kpx} y2={PADY + k.y} stroke={C.navy} strokeWidth={1.2} />); });
        const x1 = Math.min(...ends), x2 = Math.max(...ends);
        if (x2 > x1) bels.push(<line key={"h" + key} x1={x1} y1={busY} x2={x2} y2={busY} stroke={C.navy} strokeWidth={1.2} />);
      }
    }
    const cap = D.usableW;
    const svg = <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid meet" style={fit ? { width: Math.min(W, cap), maxWidth: "100%", height: "auto", display: "block", margin: "0 auto" } : { width: W, height: "auto", display: "block", margin: "0 auto" }} xmlns="http://www.w3.org/2000/svg">{bels}</svg>;
    return { svg, W };
  }

  const els = [];
  for (const n of L.flat) {
    const nx = PADX + n.x - NW / 2, ny = PADY + n.y;
    const missing = !n.kids.length && !n.collapsed && isAssemblyLike(n.part) && !purchased[n.pn];
    const isPurch = !n.kids.length && !n.collapsed && isAssemblyLike(n.part) && purchased[n.pn];
    const isTop = n.depth === 0;
    const key = n.pn + "_" + n.depth + "_" + Math.round(n.x);
    els.push(<rect key={"b" + key} x={nx} y={ny} width={NW} height={NH} rx={3}
      fill={missing ? "#FFF9E8" : "#fff"} stroke={missing ? "#B8860B" : isTop ? C.navy : n.collapsed ? C.navy2 : "#444"}
      strokeWidth={isTop ? 2 : n.collapsed ? 1.6 : 1.2} strokeDasharray={missing ? "5 3" : "none"} />);
    els.push(<text key={"p" + key} x={nx + NW / 2} y={ny + (compact ? 12 : 14)} textAnchor="middle"
      fontFamily={MONO} fontSize={fPN} fontWeight={700} fill={isTop ? C.navy : (n.kids.length || n.collapsed) ? C.navy2 : "#111"}>{n.pn}</text>);
    wrapText(n.part.desc, dChars, dLines).forEach((l, li) => els.push(
      <text key={"d" + key + li} x={nx + NW / 2} y={ny + (compact ? 21 : 26) + li * (compact ? 7 : 9)} textAnchor="middle" fontSize={fD} fill="#333">{l}</text>));
    els.push(<text key={"q" + key} x={nx + NW / 2} y={ny + NH - (compact ? 4 : 6)} textAnchor="middle" fontSize={fQ} fontWeight={600} fill="#222">{"QTY: " + (n.row ? n.row.qty : "1")}</text>);
    if (n.collapsed && sheet.refs[n.pn]) els.push(<text key={"r" + key} x={nx + NW / 2} y={ny + NH + 11} textAnchor="middle" fontSize={7.5} fontWeight={700} fill={C.navy2}>{"(SEE SHEET " + sheet.refs[n.pn] + ")"}</text>);
    if (missing) els.push(<text key={"m" + key} x={nx + NW / 2} y={ny + NH + 11} textAnchor="middle" fontSize={7.5} fontWeight={700} fill="#B8860B">▲ NO BOM — VERIFY</text>);
    if (isPurch) els.push(<text key={"u" + key} x={nx + NW / 2} y={ny + NH + 11} textAnchor="middle" fontSize={7.5} fontWeight={700} fill="#666">(PURCHASED)</text>);
    const hasChildren = (n.inline && n.inline.length) || (n.stacked && n.stacked.length);
    if (hasChildren) {
      const busY = PADY + (n.depth + 1) * ROWH - 16, px = PADX + n.x;
      els.push(<line key={"v" + key} x1={px} y1={ny + NH} x2={px} y2={busY} stroke="#444" strokeWidth={1} />);
      const ends = [px];
      n.inline.forEach(k => {
        ends.push(PADX + k.x);
        els.push(<line key={"c" + key + k.pn} x1={PADX + k.x} y1={busY} x2={PADX + k.x} y2={PADY + k.y} stroke="#444" strokeWidth={1} />);
      });
      if (n.stacked && n.stacked.length) {
        const railX = PADX + n.stackX - 8;
        ends.push(railX);
        const lastY = PADY + n.stacked[n.stacked.length - 1].y + NH / 2;
        els.push(<line key={"rl" + key} x1={railX} y1={busY} x2={railX} y2={lastY} stroke="#444" strokeWidth={1} />);
        n.stacked.forEach(k => {
          const ky = PADY + k.y + NH / 2;
          els.push(<line key={"st" + key + k.pn} x1={railX} y1={ky} x2={PADX + k.x - NW / 2} y2={ky} stroke="#444" strokeWidth={1} />);
        });
      }
      const x1 = Math.min(...ends), x2 = Math.max(...ends);
      if (x2 > x1) els.push(<line key={"h" + key} x1={x1} y1={busY} x2={x2} y2={busY} stroke="#444" strokeWidth={1} />);
    }
  }
  for (const n of callouts) {
    const cy = PADY + (n.depth + 1) * ROWH + NH + 16, cw = 128, cx = PADX + n.x - cw / 2;
    els.push(<rect key={"cb" + n.pn} x={cx} y={cy} width={cw} height={22} fill="none" stroke="#888" strokeWidth={1} strokeDasharray="5 3" rx={2} />);
    els.push(<text key={"ct" + n.pn} x={cx + cw / 2} y={cy + 14} textAnchor="middle" fontSize={7.5} fill="#555">{"(SEE " + n.pn + " BOM)"}</text>);
    els.push(<line key={"cl" + n.pn} x1={PADX + n.x} y1={PADY + (n.depth + 1) * ROWH + NH} x2={PADX + n.x} y2={cy} stroke="#888" strokeWidth={.8} strokeDasharray="3 3" />);
  }
  const cap = D.usableW;
  const svg = <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid meet" style={fit ? { width: Math.min(W, cap), maxWidth: "100%", height: "auto", display: "block", margin: "0 auto" } : { width: W, height: "auto", display: "block", margin: "0 auto" }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
  return { svg, W };
}

function TreeDoc({ bom, excluded, tops, cfgName, m, purchased, profile, customer, sheetSize }) {
  const P = activeProfile(profile);
  const [fit, setFit] = useState(false); // false = actual readable size (scroll); true = fit to width
  const D = sheetDims(sheetSize || "letter");
  // sheet plan across all tops
  let allSheets = [];
  tops.forEach(t => { allSheets = allSheets.concat(planSheets(bom, excluded, t, undefined, D)); });
  allSheets.forEach(s => { s.D = D; });
  allSheets.forEach((s, i) => s.sheetNo = i + 1);
  const byTop = {}; allSheets.forEach(s => { if (byTop[s.top] === undefined) byTop[s.top] = s.sheetNo; });
  allSheets.forEach(s => { s.refs = {}; s.collapse.forEach(pn => { if (byTop[pn]) s.refs[pn] = byTop[pn]; }); });
  const nSheets = allSheets.length;

  const blocks = tops.map(top => scopedTree(bom, excluded, top));
  const flat = blocks.flat();
  const th = { background: C.navy, color: "#fff", padding: "4px 6px", textAlign: "left", fontSize: 9.5 };
  const td = { border: `1px solid ${C.line}`, padding: "3px 6px", verticalAlign: "top" };
  const missing = flat.filter(r => !r.isAsm && isAssemblyLike(r.part) && !purchased[r.part.pn]);
  const tbC = { padding: "3px 8px", borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}`, fontSize: 9 };
  const oneSheet = D.oneSheet;
  return (
    <Sheet wide={oneSheet}>
      {/* ---- drawing title block ---- */}
      <div style={{ border: `1.5px solid #111`, marginBottom: oneSheet ? 6 : 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "200px 1fr 250px" }}>
          <div style={{ padding: "8px 10px", borderRight: "1px solid #111" }}>
            {P.id === "island"
              ? <img src={ISLAND_LOGO} alt="Island Components" style={{ height: 34, display: "block" }} />
              : <div style={{ fontWeight: 900, fontSize: 20, color: C.navy, fontStyle: "italic", letterSpacing: "-.02em" }}>EZ<span style={{ fontSize: 13, fontStyle: "normal", letterSpacing: ".08em" }}>MOTORS</span></div>}
            <div style={{ fontSize: 8.5, marginTop: 2 }}>{P.company}</div>
            <div style={{ fontSize: 7.5, color: "#555" }}>{P.address}</div>
            <div style={{ fontSize: 7.5, color: "#555" }}>{P.cage}</div>
          </div>
          <div style={{ padding: "8px 10px", textAlign: "center", borderRight: "1px solid #111" }}>
            <div style={{ fontWeight: 800, fontSize: 16 }}>{(bom.parts[tops[0]] && bom.parts[tops[0]].desc ? bom.parts[tops[0]].desc : tops[0]).toUpperCase()}</div>
            <div style={{ fontSize: 13, letterSpacing: ".1em", marginTop: 2 }}>FAMILY TREE</div>
            <div style={{ fontSize: 10, marginTop: 3 }}>TOP LEVEL ASSEMBL{tops.length > 1 ? "IES" : "Y"}: <b style={{ fontFamily: MONO }}>{tops.join(" + ")}</b> — {cfgName}</div>
          </div>
          <div style={{ fontSize: 8.5 }}>
            {[["DOCUMENT NO.", P.docNo.F(tops.join("+")), "REV."], ["DATE", m.date, m.rev || "1"], ["DRAWN BY:", "Engineering", nSheets + " SHEET" + (nSheets > 1 ? "S" : "")], ["CHECKED BY:", "Quality", ""], ["APPROVED BY:", "", ""]].map((r, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "82px 1fr 60px" }}>
                <div style={{ ...tbC, fontWeight: 700 }}>{r[0]}</div>
                <div style={{ ...tbC, fontFamily: MONO }}>{r[1]}</div>
                <div style={{ ...tbC, borderRight: "none", textAlign: "center", fontWeight: i === 0 ? 700 : 400 }}>{r[2]}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ---- drawing sheets ---- */}
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 4, gap: 6 }}>
        <span style={{ fontSize: 10, color: "#999", alignSelf: "center" }}>{oneSheet ? "Large-format sheet — use Actual size for legibility, Fit width for overview" : ""}</span>
        <button onClick={() => setFit(!fit)} style={{ border: `1px solid ${C.line}`, background: "#fff", fontSize: 10, padding: "3px 8px", cursor: "pointer", color: "#666" }}>
          {fit ? "⤢ Actual size (scroll)" : "⤡ Fit width"}
        </button>
      </div>
      {allSheets.map(sh => {
        const d = SheetDrawing({ bom, sheet: sh, purchased, fit: oneSheet ? true : fit, qaField: P.qaField });
        const shTop = bom.parts[sh.top] || {};
        return (
          <div key={sh.sheetNo} style={{ border: "1px solid #111", marginBottom: oneSheet ? 4 : 10, background: "#fff", breakInside: "avoid", pageBreakInside: "avoid" }}>
            <div style={{ borderBottom: "1px solid #111", padding: "3px 8px", fontSize: 9, display: "flex", justifyContent: "space-between", background: "#FAFAF8" }}>
              <span style={{ fontWeight: 700 }}>SHEET {sh.sheetNo} OF {nSheets}{sh.sheetNo > 1 ? ` — SUBASSEMBLY: ${sh.top}` : ""}{oneSheet ? " — 11×17" : ""}</span>
              <span style={{ fontFamily: MONO, color: "#666" }}>{sh.sheetNo === 1 ? tops.join(" + ") : (shTop.desc || "").toUpperCase()}</span>
            </div>
            <div style={{ padding: 8, overflowX: fit ? "hidden" : "auto", display: "flex", justifyContent: "center" }}>
              {fit ? d.svg : <div style={{ width: d.W, minWidth: d.W }}>{d.svg}</div>}
            </div>
          </div>
        );
      })}

      {/* ---- legend / revision / notes band ---- */}
      <div style={{ display: "grid", gridTemplateColumns: "150px 1fr 250px", gap: 8, marginBottom: 6, fontSize: oneSheet ? 7.5 : 8 }}>
        <div style={{ border: "1px solid #111", padding: 8, fontSize: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>LEGEND</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}><span style={{ width: 24, height: 12, border: "1.2px solid #444", display: "inline-block" }} /> = ASSEMBLY / PART</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}><span style={{ width: 24, height: 12, border: `1.4px solid ${C.navy2}`, display: "inline-block" }} /> = SEE INDICATED SHEET</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}><span style={{ width: 24, height: 12, border: "1px dashed #888", display: "inline-block" }} /> = SEE INDICATED BOM</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}><span style={{ width: 24, height: 12, border: "1.2px dashed #B8860B", background: "#FFF9E8", display: "inline-block" }} /> = NO BOM LOADED</div>
        </div>
        <div style={{ border: "1px solid #111", fontSize: 8.5 }}>
          <div style={{ textAlign: "center", fontWeight: 700, borderBottom: "1px solid #111", padding: 3 }}>REVISION HISTORY</div>
          <div style={{ display: "grid", gridTemplateColumns: "40px 70px 1fr 70px 40px 70px", fontWeight: 700, borderBottom: "1px solid #111" }}>
            {["REV.", "ECO NO.", "DESCRIPTION", "DATE", "BY", "APPROVED"].map(h => <div key={h} style={{ padding: 3, borderRight: "1px solid #ccc" }}>{h}</div>)}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "40px 70px 1fr 70px 40px 70px" }}>
            {[m.rev || "1", m.eco || "ECO-0001", m.change || "INITIAL RELEASE (GENERATED)", m.date, "ENG", ""].map((v2, i) => <div key={i} style={{ padding: 3, borderRight: "1px solid #ccc", fontFamily: MONO }}>{v2}</div>)}
          </div>
        </div>
        <div style={{ border: "1px solid #111", padding: 8, fontSize: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>NOTES:</div>
          <div>1. THIS FAMILY TREE DEPICTS THE ASSEMBLY HIERARCHY FOR {tops.join(" + ")}.</div>
          <div>2. QUANTITIES ARE FOR ONE (1) TOP LEVEL ASSEMBLY UNLESS OTHERWISE NOTED.</div>
          <div>3. ASSEMBLIES MARKED "SEE SHEET n" ARE DETAILED ON THE INDICATED SHEET OF THIS DOCUMENT.</div>
          <div>4. REFER TO PARTS LIST P-{tops.join("+")} FOR THE COMPLETE INDENTURED PARTS LIST.</div>
          <div>5. REFER TO INDIVIDUAL BOM DOCUMENTS FOR COMPLETE PART DETAILS, MATERIAL, FINISH, AND PROCUREMENT INFORMATION.</div>
          {missing.length > 0 && <div style={{ color: "#B8860B", fontWeight: 700 }}>6. ▲ {missing.length} ASSEMBL{missing.length > 1 ? "IES" : "Y"} SHOWN WITHOUT LOADED BOM — VERIFY MISSING FILE OR MARK PURCHASED: {missing.map(r => r.part.pn).join(", ")}.</div>}
        </div>
      </div>

      <div style={{ textAlign: "center", borderTop: "1.5px solid #111", marginTop: 14, paddingTop: 5, fontSize: 9.5, fontWeight: 700, letterSpacing: ".05em" }}>
        PROPRIETARY AND CONFIDENTIAL — {P.company} <span style={{ float: "right", fontWeight: 400 }}>AS9100D</span>
      </div>
    </Sheet>
  );
}


/* WI vignette library — schematic placeholder renderings per operation type.
   Line-art style: grays + navy accents. Each vignette: 300x150 viewBox SVG inner markup. */
const vN = "#1F3864", vN2 = "#2E5395", vG = "#666", vLG = "#aaa", vAMB = "#B8860B";
const S = (d, c = vG, w = 1.6, extra = "") => `<path d="${d}" fill="none" stroke="${c}" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" ${extra}/>`;
const R = (x, y, w, h, c = vG, sw = 1.6, fill = "none", rx = 2) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}" fill="${fill}" stroke="${c}" stroke-width="${sw}"/>`;
const Ci = (x, y, r, c = vG, sw = 1.6, fill = "none") => `<circle cx="${x}" cy="${y}" r="${r}" fill="${fill}" stroke="${c}" stroke-width="${sw}"/>`;
const T = (x, y, t, s = 8, c = vG, anchor = "middle", w = 400) => `<text x="${x}" y="${y}" font-size="${s}" fill="${c}" text-anchor="${anchor}" font-family="Segoe UI,Arial" font-weight="${w}">${t}</text>`;
const arrow = (x1, y1, x2, y2, c = vN2) => S(`M${x1} ${y1} L${x2} ${y2}`, c, 1.4) + S(`M${x2} ${y2} l-5 -3 M${x2} ${y2} l-3 5`, c, 1.4);

// stator cross-section (end view): OD ring, bore, slots
function statorSection(cx, cy, rOD, extra = "") {
  const rB = rOD * 0.45;
  let s = Ci(cx, cy, rOD, vG, 2) + Ci(cx, cy, rB, vG, 1.6);
  for (let i = 0; i < 12; i++) {
    const a = i * Math.PI / 6;
    const x1 = cx + Math.cos(a) * rB, y1 = cy + Math.sin(a) * rB;
    const x2 = cx + Math.cos(a) * (rOD * 0.82), y2 = cy + Math.sin(a) * (rOD * 0.82);
    s += S(`M${x1} ${y1} L${x2} ${y2}`, vLG, 1.2);
  }
  return s + extra;
}
// rotor side view: shaft with magnet band
function rotorSide(x, y, len) {
  return R(x, y + 10, len, 14, vG, 1.8) +                       // shaft
    R(x + len * 0.25, y, len * 0.5, 34, vN2, 1.8, "#EAF0FA") +  // magnet band
    S(`M${x + len * 0.25} ${y + 8} h${len * 0.5} M${x + len * 0.25} ${y + 17} h${len * 0.5} M${x + len * 0.25} ${y + 26} h${len * 0.5}`, vN2, 1);
}
const V = {};

V.kitting = { label: "Kitting tray, BOM check sheet, labeled bins", svg:
  R(20, 40, 130, 80, vG, 2) + S("M20 66 h130 M20 92 h130 M63 40 v80 M107 40 v80", vLG, 1.2) +
  Ci(41, 53, 6, vN2) + R(78, 50, 18, 8, vN2) + Ci(128, 79, 8, vG) + R(30, 100, 22, 12, vG) + R(75, 76, 20, 10, vN2) +
  R(180, 35, 95, 90, vG, 1.8, "#fff") + S("M190 52 h60 M190 66 h75 M190 80 h68 M190 94 h75 M190 108 h55", vLG, 1.6) +
  S("M184 50 l4 4 l7 -8", vN2, 2) + S("M184 64 l4 4 l7 -8", vN2, 2) + S("M184 78 l4 4 l7 -8", vN2, 2) +
  T(85, 133, "KITTED PARTS", 8) + T(227, 133, "BOM VERIFICATION", 8) },

V.winding = { label: "Coil winder with tensioner and turns counter", svg:
  Ci(70, 80, 34, vG, 2) + Ci(70, 80, 12, vG, 1.6) +
  S("M70 46 a34 34 0 0 1 24 10", vN2, 3) + S("M70 52 a28 28 0 0 1 20 8", vN2, 2) +
  S("M104 66 C 140 40, 170 40, 205 55", vN2, 1.6) +
  R(200, 48, 26, 16, vG, 1.8) + T(213, 59, "T", 9, vG) +
  S("M226 56 C 250 60, 258 70, 262 84", vN2, 1.6) +
  R(248, 84, 32, 26, vG, 1.8, "#fff") + T(264, 101, "0248", 9, vN, "middle", 700) + T(264, 122, "TURNS", 7) +
  T(70, 133, "WINDING FORM", 8) + T(213, 40, "TENSIONER", 7) },

V.insertCoils = { label: "Coil insertion into slots with liners and wedges", svg:
  statorSection(90, 82, 52) +
  S("M90 30 C 120 18, 150 22, 168 40", vN2, 3) +
  S("M172 44 l-14 10", vN2, 2) + arrow(150, 20, 118, 42) +
  R(196, 46, 84, 70, "#fff", 0) +
  R(200, 50, 76, 62, vG, 1.6) + T(238, 62, "SLOT DETAIL", 7) +
  S("M212 70 v34 M264 70 v34", vG, 1.6) + S("M214 72 v30 M262 72 v30", vLG, 1) +
  Ci(230, 88, 5, vN2) + Ci(246, 88, 5, vN2) + Ci(238, 98, 5, vN2) +
  S("M212 70 h52", vAMB, 2.5) + T(238, 120, "LINER · COILS · WEDGE", 7) +
  T(90, 145, "INSERT PER SEQUENCE", 8) },

V.rotationCheck = { label: "Phase sequence / electrical rotation verification", svg:
  statorSection(78, 82, 46) +
  S("M78 128 v14 M60 126 v16 M96 126 v16", vN2, 2) +
  T(60, 150, "A", 7, vN2) + T(78, 150, "B", 7, vN2) + T(96, 150, "C", 7, vN2) +
  R(180, 48, 96, 66, vG, 1.8, "#fff") +
  S("M192 96 a20 14 0 0 1 72 0", vN2, 2) + arrow(258, 92, 264, 78) +
  T(228, 74, "A→B→C", 10, vN, "middle", 700) + T(228, 108, "ROTATION OK", 7) +
  S("M124 100 C 150 110, 160 100, 180 92", vLG, 1.4, 'stroke-dasharray="4 3"') },

V.leads = { label: "Phase joint: crimp/solder, sleeve, phase ID", svg:
  S("M20 70 h70", vN2, 4) + S("M195 70 h85", vN2, 4) +
  R(90, 60, 105, 20, vG, 1.8, "#EAF0FA", 8) + S("M104 60 v20 M181 60 v20", vG, 1.2) +
  T(142, 74, "SLEEVE", 8, vN) +
  R(60, 44, 22, 12, vAMB, 1.6, "#FFF6E0") + T(71, 53, "A", 8, vAMB, "middle", 700) +
  arrow(142, 40, 142, 56) + T(142, 34, "INSULATED JOINT", 7) +
  S("M225 88 q10 14 30 14", vG, 1.4) + R(252, 96, 30, 14, vG, 1.4) + T(267, 106, "S/R", 7) +
  T(150, 132, "STRAIN RELIEF PER DRAWING", 8) },

V.lacing = { label: "End-turn lacing pattern", svg:
  S("M30 95 h240", vG, 2) +
  S("M40 95 a14 22 0 0 1 28 0 M68 95 a14 22 0 0 1 28 0 M96 95 a14 22 0 0 1 28 0 M124 95 a14 22 0 0 1 28 0 M152 95 a14 22 0 0 1 28 0 M180 95 a14 22 0 0 1 28 0 M208 95 a14 22 0 0 1 28 0", vN2, 2.4) +
  S("M40 84 q14 -20 28 0 q14 -20 28 0 q14 -20 28 0 q14 -20 28 0 q14 -20 28 0 q14 -20 28 0 q14 -20 28 0", vAMB, 1.6) +
  S("M40 84 l0 0 M236 84 l0 0", vAMB, 1.6) +
  T(150, 128, "UNIFORM LACING — SECURE ALL END TURNS", 8) },

V.forming = { label: "End-turn forming with envelope gauge", svg:
  S("M60 110 h180", vG, 2.4) +
  S("M80 110 a35 30 0 0 1 140 0", vN2, 3) +
  S("M70 110 a40 38 0 0 1 160 0", vAMB, 1.6, 'stroke-dasharray="6 4"') +
  T(150, 58, "GAUGE ENVELOPE", 7, vAMB) +
  R(130, 26, 40, 14, vG, 1.6) + T(150, 36, "FORM", 7) + arrow(150, 42, 150, 62) +
  T(150, 134, "FORM TO DRAWING ENVELOPE — VERIFY CLEARANCE", 8) },

V.electricalTest = { label: "Electrical test station: IR / resistance / surge / hipot", svg:
  R(30, 45, 90, 75, vG, 1.8) + statorSection(75, 82, 30) +
  S("M120 60 h40 M120 100 h40", vN2, 1.6) +
  R(160, 40, 116, 84, vG, 1.8, "#fff") +
  R(170, 50, 96, 26, vN, 1.4, "#EAF0FA") + T(218, 67, "12.48 mΩ", 10, vN, "middle", 700) +
  S("M170 88 h96 M170 100 h96 M170 112 h60", vLG, 1.4) +
  S("M172 86 l4 4 l6 -7", vN2, 1.8) + S("M172 98 l4 4 l6 -7", vN2, 1.8) +
  T(218, 135, "RECORD ALL VALUES + EQUIPMENT ID", 7.5) },

V.impregnation = { label: "Varnish impregnation and cure oven", svg:
  R(28, 60, 100, 62, vG, 2) + S("M28 76 h100", vN2, 1.6) +
  S("M36 76 q6 -6 12 0 t12 0 t12 0 t12 0 t12 0 t12 0 t12 0", vN2, 1.4) +
  statorSection(78, 100, 22) + T(78, 136, "VARNISH TANK", 7.5) +
  R(168, 48, 108, 74, vG, 2) + S("M168 62 h108", vG, 1.2) + Ci(258, 55, 4, vG) +
  S("M186 80 q4 -10 0 -18 M202 84 q4 -10 0 -18 M218 80 q4 -10 0 -18", vAMB, 1.8) +
  T(222, 106, "CURE PER SPEC", 7.5) + T(222, 135, "RECORD CYCLE + CHART", 7.5) },

V.visual = { label: "Manufacturing visual inspection under magnification", svg:
  statorSection(90, 88, 44) +
  Ci(190, 66, 30, vN, 2.4) + S("M212 88 L246 122", vN, 5) +
  Ci(190, 66, 22, vLG, 1) + S("M180 58 q10 -8 18 0", vN2, 1.6) +
  T(190, 140, "NO VOIDS · NO LOOSE WIRES · CLEAN BORE", 8) },

V.cleanPrep = { label: "Bond surface preparation: solvent wipe, masked datums", svg:
  R(30, 74, 190, 16, vG, 1.8) +
  R(30, 70, 30, 24, vAMB, 1.6, "#FFF6E0") + R(190, 70, 30, 24, vAMB, 1.6, "#FFF6E0") +
  T(45, 64, "MASK", 6.5, vAMB) + T(205, 64, "MASK", 6.5, vAMB) +
  R(100, 34, 44, 26, vG, 1.6, "#fff", 4) + S("M104 60 l10 12 h16 l10 -12", vG, 1.4) +
  arrow(122, 74, 122, 70) + T(122, 28, "LINT-FREE WIPE", 7) +
  S("M244 60 l6 18 l8 -4", vN2, 2) + T(258, 96, "ONE PASS,", 7) + T(258, 105, "ONE USE", 7) +
  T(125, 130, "CLEAN UNTIL RESIDUE-FREE — NO BARE-HAND CONTACT", 7.5) },

V.polarity = { label: "Magnet dry-fit in fixture with gauss polarity check", svg:
  Ci(90, 84, 44, vG, 2) + Ci(90, 84, 14, vG, 1.6) +
  (() => { let s = ""; for (let i = 0; i < 8; i++) { const a = i * Math.PI / 4 - Math.PI / 8;
    const x = 90 + Math.cos(a) * 29, y = 84 + Math.sin(a) * 29;
    s += R(x - 9, y - 6, 18, 12, i % 2 ? vN2 : vAMB, 1.4, i % 2 ? "#EAF0FA" : "#FFF6E0", 2) +
      T(x, y + 3, i % 2 ? "S" : "vN", 7, i % 2 ? vN2 : vAMB, "middle", 700); } return s; })() +
  R(190, 56, 86, 56, vG, 1.8, "#fff") + T(233, 78, "vN-S-vN-S", 10, vN, "middle", 700) +
  T(233, 96, "GAUSS CHECK", 7) + S("M134 84 h50", vLG, 1.4, 'stroke-dasharray="4 3"') +
  T(90, 142, "MAP EACH POSITION ON TRAVELER", 7.5) },

V.magnetBond = { label: "Adhesive application pattern and magnet seating", svg:
  rotorSide(40, 60, 200) +
  R(120, 34, 40, 18, vN2, 1.6, "#EAF0FA") + T(140, 47, "MAG", 7, vN2) +
  arrow(140, 54, 140, 60) +
  Ci(112, 42, 3, vAMB, 1, vAMB) + Ci(122, 30, 3, vAMB, 1, vAMB) + Ci(158, 30, 3, vAMB, 1, vAMB) +
  T(210, 40, "ADHESIVE PATTERN", 7, vAMB) + T(210, 50, "PER RELEASED PROCESS", 6.5, vAMB) +
  T(140, 130, "SEAT FULLY — NO ADHESIVE ON POLES OR DATUMS", 7.5) },

V.cure = { label: "Cure fixture in oven with chart recorder", svg:
  R(60, 40, 180, 88, vG, 2) + S("M60 56 h180", vG, 1.2) + Ci(226, 48, 4, vG) +
  rotorSide(95, 78, 110) + R(88, 72, 124, 48, vN2, 1.4, "none", 4) +
  T(150, 66, "CURE FIXTURE", 7, vN2) +
  R(252, 60, 34, 56, vG, 1.6, "#fff") + S("M256 106 q8 -26 6 -40 q8 24 6 38 q8 -20 6 -34", vN2, 1.2) +
  T(269, 128, "CHART", 7) + T(150, 142, "TIME / TEMP PER RELEASED SCHEDULE", 7.5) },

V.grind = { label: "Cylindrical grinding between centers, coolant + extraction", svg:
  rotorSide(30, 76, 150) +
  S("M22 93 l8 -6 l0 12 z M188 87 l-8 6 l8 6 z", vG, 1.6) +
  Ci(212, 62, 34, vG, 2.4) + Ci(212, 62, 6, vG, 1.4) +
  S("M212 28 a34 34 0 0 1 20 8", vLG, 2) +
  S("M186 84 q-6 10 -14 12 M192 90 q-4 10 -10 14", vN2, 1.4) +
  R(238, 100, 44, 26, vG, 1.6) + T(260, 117, "EXTRACT", 6.5) +
  T(120, 138, "LIGHT PASSES — NO THERMAL DAMAGE", 7.5) },

V.sleeve = { label: "Retention sleeve installation with alignment fixture", svg:
  rotorSide(40, 66, 190) +
  R(96, 52, 118, 46, vN, 2, "none", 6) + T(155, 46, "SLEEVE", 7.5, vN) +
  arrow(155, 30, 155, 50) + T(155, 24, "CONTROLLED PRESS / THERMAL", 7) +
  Ci(100, 108, 3, vAMB, 1, vAMB) + Ci(210, 108, 3, vAMB, 1, vAMB) +
  T(155, 122, "UNIFORM SQUEEZE-OUT AT EDGES", 7, vAMB) +
  T(155, 142, "MAINTAIN AXIAL POSITION PER DRAWING", 7.5) },

V.inspect = { label: "Inspection: V-blocks, indicator, calibrated instruments", svg:
  S("M60 116 l22 -22 l22 22 z M176 116 l22 -22 l22 22 z", vG, 2) +
  rotorSide(58, 62, 168) + S("M30 116 h240", vG, 2) +
  R(128, 22, 26, 20, vG, 1.8, "#fff") + Ci(141, 32, 7, vG, 1.2) + S("M141 32 l4 -5", vN2, 1.6) +
  S("M141 42 v20", vG, 1.8) + S("M137 62 h8", vG, 1.8) +
  T(141, 16, "INDICATOR", 7) + T(150, 138, "RUNOUT · SIZE · FINISH PER DRAWING", 7.5) },

V.rotorInsert = { label: "Guided rotor insertion tooling — controlled magnetic pull-in", svg:
  R(150, 44, 110, 76, vG, 2) + Ci(205, 82, 26, vG, 1.6) +
  rotorSide(30, 68, 110) +
  R(24, 58, 124, 50, vN2, 1.6, "none", 6) + T(86, 52, "GUIDE TOOL", 7, vN2) +
  arrow(150, 82, 172, 82) +
  T(205, 134, "NEVER INSERT BY HAND", 7.5, vAMB, "middle", 700) },

V.statorInstall = { label: "Stator installation into housing — heat / press / bond, lead clocking", svg:
  R(170, 36, 100, 92, vG, 2.4) + Ci(220, 82, 34, vLG, 1.6) +
  statorSection(80, 82, 40) +
  arrow(126, 82, 168, 82) +
  S("M186 30 q4 -10 0 -16 M200 32 q4 -10 0 -16 M214 30 q4 -10 0 -16", vAMB, 1.8) +
  T(200, 12, "HEAT PER PROCESS", 6.5, vAMB) +
  S("M80 42 v-14", vN2, 2) + T(80, 20, "LEAD CLOCKING", 6.5, vN2) +
  T(150, 144, "SEAT TO DATUM — PROTECT WINDINGS", 7.5) },

V.torque = { label: "Cross-pattern torque sequence with calibrated wrench", svg:
  Ci(100, 82, 44, vG, 2) + Ci(100, 82, 16, vLG, 1.4) +
  (() => { let s = ""; const seq = [1, 4, 2, 5, 3, 6]; for (let i = 0; i < 6; i++) { const a = i * Math.PI / 3 - Math.PI / 2;
    const x = 100 + Math.cos(a) * 32, y = 82 + Math.sin(a) * 32;
    s += Ci(x, y, 6, vN2, 1.6) + T(x, y + 3, String(seq[i]), 7, vN, "middle", 700); } return s; })() +
  S("M170 60 h70 l14 10 l-14 10 h-70 z", vG, 1.8) + R(160, 62, 14, 16, vG, 1.8) +
  T(205, 100, "CAL. TORQUE WRENCH", 7) + T(150, 138, "TORQUE IN SEQUENCE — WITNESS MARK", 7.5) },

V.stack = { label: "Lamination stacking fixture with alignment key", svg:
  R(90, 100, 120, 12, vG, 2) + S("M96 100 v-56 M204 100 v-56", vG, 1.8) +
  (() => { let s = ""; for (let i = 0; i < 8; i++) s += R(104, 92 - i * 7, 92, 5, i % 2 ? vLG : vG, 1); return s; })() +
  R(146, 30, 8, 66, vN2, 1.6, "#EAF0FA") + T(150, 24, "KEY", 6.5, vN2) +
  arrow(150, 8, 150, 18) +
  T(150, 132, "STACK TO COUNT / HEIGHT — MAINTAIN SLOT ALIGNMENT", 7.5) },

V.press = { label: "Arbor press — force on fitted ring only", svg:
  R(120, 20, 60, 14, vG, 2) + S("M150 34 v34", vG, 4) + R(132, 68, 36, 10, vG, 2) +
  Ci(150, 104, 24, vG, 2) + Ci(150, 104, 12, vLG, 1.6) +
  arrow(150, 46, 150, 62) +
  T(226, 100, "FORCE ON", 7) + T(226, 109, "FITTED RING", 7) +
  T(150, 142, "NEVER THROUGH ROLLING ELEMENTS", 7.5, vAMB) },

V.gear = { label: "Planetary gear train assembly with timing marks", svg:
  Ci(150, 80, 48, vG, 2) + Ci(150, 80, 12, vN2, 1.8) +
  Ci(150, 46, 14, vG, 1.8) + Ci(120, 97, 14, vG, 1.8) + Ci(180, 97, 14, vG, 1.8) +
  Ci(150, 46, 3, vAMB, 1, vAMB) + Ci(120, 97, 3, vAMB, 1, vAMB) + Ci(180, 97, 3, vAMB, 1, vAMB) +
  T(228, 60, "TIMING", 7, vAMB) + T(228, 70, "MARKS", 7, vAMB) +
  T(150, 144, "VERIFY FREE ROTATION — LUBE PER SPEC", 7.5) },

V.stock = { label: "Preservation, identification, and protective packaging", svg:
  S("M70 60 l60 -22 l60 22 v50 l-60 22 l-60 -22 z", vG, 2) + S("M70 60 l60 22 l60 -22 M130 82 v50", vG, 1.6) +
  R(150, 66, 34, 20, vN2, 1.4, "#fff") + S("M154 72 h26 M154 78 h18", vN2, 1.2) +
  T(230, 56, "LOT / SERIAL", 7) + T(230, 66, "NHA REF", 7) +
  S("M60 44 a10 10 0 0 1 14 -8", vAMB, 1.8) + T(52, 30, "CAPS /", 6.5, vAMB) + T(52, 39, "PROTECT", 6.5, vAMB) +
  T(150, 144, "STOCK TRANSACTION TO NHA", 7.5) },

V.assembly = { label: "Assembly operation per released drawing", svg:
  Ci(110, 82, 36, vG, 2) + Ci(110, 82, 14, vLG, 1.6) +
  S("M170 60 h60 l12 9 l-12 9 h-60 z", vG, 1.8) +
  R(160, 62, 12, 14, vG, 1.8) +
  T(150, 138, "ASSEMBLE PER DRAWING — RECORD TORQUES", 7.5) };

const PICK = [
  [/stator.*housing|housing.*stator/i, "statorInstall"],
  [/pre-varnish/i, "electricalTest"],
  [/post-grind/i, "inspect"],
  [/kit/i, "kitting"],
  [/wind/i, "winding"],
  [/insert coil/i, "insertCoils"],
  [/rotation/i, "rotationCheck"],
  [/lead|terminate|connector/i, "leads"],
  [/lace/i, "lacing"],
  [/form/i, "forming"],
  [/varnish|impregnat/i, "impregnation"],
  [/clean up|visual inspection \(man/i, "visual"],
  [/bond surface|preparation/i, "cleanPrep"],
  [/dry-fit|polarity/i, "polarity"],
  [/magnet bonding/i, "magnetBond"],
  [/cure/i, "cure"],
  [/grind/i, "grind"],
  [/sleeve/i, "sleeve"],
  [/rotor & bearing|install rotor/i, "rotorInsert"],
  [/endplay|close|air gap/i, "torque"],
  [/stack/i, "stack"],
  [/press bearing/i, "press"],
  [/gear/i, "gear"],
  [/stock|preserve|package|ship/i, "stock"],
  [/machin/i, "grind"],
  [/test|verification|retest|electrical/i, "electricalTest"],
  [/inspect/i, "inspect"],
];
function pickVignette(title) {
  for (const [re, k] of PICK) if (re.test(title)) return { key: k, ...V[k] };
  return { key: "assembly", ...V.assembly };
}
function vignetteSvg(title) {
  const v = pickVignette(title);
  return { label: v.label, svg: `<svg viewBox="0 0 300 155" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:340px;height:auto;display:block;margin:0 auto">${v.svg}</svg>` };
}


/* ---- Parts List document (P-<top PN>) ---- */
function PartsListDoc({ bom, excluded, tops, cfgName, m, purchased, profile, customer }) {
  const P = activeProfile(profile);
  const topPart = bom.parts[tops[0]] || {};
  const flat = tops.flatMap(t => scopedTree(bom, excluded, t));
  const missing = flat.filter(r => !r.isAsm && isAssemblyLike(r.part) && !purchased[r.part.pn]);
  const th = { background: C.navy, color: "#fff", padding: "4px 6px", textAlign: "left", fontSize: 9 };
  const td = { border: `1px solid ${C.line}`, padding: "3px 6px", verticalAlign: "top", fontSize: 10 };
  const tbC = { padding: "3px 8px", borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}`, fontSize: 9 };
  return (
    <Sheet>
      <div style={{ border: `1.5px solid #111`, marginBottom: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "200px 1fr 250px" }}>
          <div style={{ padding: "8px 10px", borderRight: "1px solid #111" }}>
            {P.id === "island"
              ? <img src={ISLAND_LOGO} alt="Island Components" style={{ height: 34, display: "block" }} />
              : <div style={{ fontWeight: 900, fontSize: 20, color: C.navy, fontStyle: "italic", letterSpacing: "-.02em" }}>EZ<span style={{ fontSize: 13, fontStyle: "normal", letterSpacing: ".08em" }}>MOTORS</span></div>}
            <div style={{ fontSize: 8.5, marginTop: 2 }}>{P.company}</div>
            <div style={{ fontSize: 7.5, color: "#555" }}>{P.address}</div>
            <div style={{ fontSize: 7.5, color: "#555" }}>{P.cage}</div>
          </div>
          <div style={{ padding: "8px 10px", textAlign: "center", borderRight: "1px solid #111" }}>
            <div style={{ fontWeight: 800, fontSize: 16 }}>{(topPart.desc || tops[0]).toUpperCase()}</div>
            <div style={{ fontSize: 13, letterSpacing: ".1em", marginTop: 2 }}>PARTS LIST</div>
            <div style={{ fontSize: 10, marginTop: 3 }}>TOP LEVEL ASSEMBL{tops.length > 1 ? "IES" : "Y"}: <b style={{ fontFamily: MONO }}>{tops.join(" + ")}</b> — {cfgName} · SEE FAMILY TREE F-{tops.join("+")}</div>
          </div>
          <div style={{ fontSize: 8.5 }}>
            {[["DOCUMENT NO.", P.docNo.P(tops.join("+")), "REV."], ["DATE", m.date, m.rev || "1"], ["DRAWN BY:", "Engineering", "1 OF 1"], ["CHECKED BY:", "Quality", ""], ["APPROVED BY:", "", ""]].map((r, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "82px 1fr 60px" }}>
                <div style={{ ...tbC, fontWeight: 700 }}>{r[0]}</div>
                <div style={{ ...tbC, fontFamily: MONO }}>{r[1]}</div>
                <div style={{ ...tbC, borderRight: "none", textAlign: "center", fontWeight: i === 0 ? 700 : 400 }}>{r[2]}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr>{["Lvl", "Find", "Qty", "UOM", "Part Number", "Rev", "Description", "Make/Buy", "Material", "Remarks", "Status"].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
        <tbody>
          {flat.map((r, i) => {
            const bg = r.depth === 0 ? C.ltblue : r.isAsm ? C.gray : "transparent";
            const miss = !r.isAsm && isAssemblyLike(r.part) && !purchased[r.part.pn];
            const purch = !r.isAsm && isAssemblyLike(r.part) && purchased[r.part.pn];
            return (
              <tr key={i} style={{ background: miss ? "#FFF9E8" : bg, fontWeight: r.depth === 0 || r.isAsm ? 700 : 400 }}>
                <td style={{ ...td, textAlign: "center" }}>{r.depth}</td>
                <td style={{ ...td, textAlign: "center", fontFamily: MONO }}>{r.row ? r.row.find : "—"}</td>
                <td style={{ ...td, textAlign: "center" }}>{r.row ? r.row.qty : "1"}</td>
                <td style={{ ...td, textAlign: "center" }}>{r.row ? r.row.uom : "EA"}</td>
                <td style={{ ...td, fontFamily: MONO, paddingLeft: 6 + r.depth * 12 }}>{r.part.pn}</td>
                <td style={{ ...td, textAlign: "center" }}>{r.part.rev}</td>
                <td style={td}>{r.part.desc}</td>
                <td style={{ ...td, textAlign: "center" }}>{r.part.mb || "—"}</td>
                <td style={{ ...td, fontSize: 9 }}>{r.part.mat || ""}</td>
                <td style={{ ...td, fontSize: 9 }}>{r.part.rem || ""}</td>
                <td style={{ ...td, fontSize: 8.5, color: miss ? "#B8860B" : purch ? "#666" : "#999", fontWeight: miss ? 700 : 400 }}>{r.isAsm ? "BOM \u2713" : miss ? "\u25B2 NO BOM" : purch ? "PURCHASED" : ""}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {missing.length > 0 && <div style={{ marginTop: 8, fontSize: 9.5, color: "#B8860B", fontWeight: 700 }}>▲ {missing.length} ASSEMBL{missing.length > 1 ? "IES" : "Y"} LISTED WITHOUT LOADED BOM — VERIFY MISSING FILE OR CONFIRM PURCHASED COMPLETE.</div>}
      <div style={{ marginTop: 10, fontSize: 8.5, color: "#555" }}>NOTES: 1. LEVEL 0 = TOP LEVEL / END ITEM. QUANTITIES ARE PER ONE UNIT OF THE NEXT HIGHER ASSEMBLY. 2. THIS PARTS LIST IS GENERATED FROM THE RELEASED BOM STRUCTURE AND IS SUBORDINATE TO THE RELEASED BOM DOCUMENTS.</div>
      <div style={{ textAlign: "center", borderTop: "1.5px solid #111", marginTop: 12, paddingTop: 5, fontSize: 9.5, fontWeight: 700, letterSpacing: ".05em" }}>
        PROPRIETARY AND CONFIDENTIAL — {P.company} <span style={{ float: "right", fontWeight: 400 }}>AS9100D</span>
      </div>
    </Sheet>
  );
}

/* ---- Traveler docs ---- */
function IslandBrandBlock({ sub }) {
  return (
    <div style={{ textAlign: "center", marginBottom: 6 }}>
      <img src={ISLAND_LOGO} alt="Island Components" style={{ height: 52, display: "block", margin: "0 auto 3px" }} />
      <div style={{ fontSize: 8, color: "#555" }}>ISLAND COMPONENTS GROUP, INC. · 210 Marcus Blvd. Hauppauge, N.Y. 11788 · Tel (631) 563-4224</div>
      {sub && <div style={{ fontSize: 8, color: "#777" }}>{sub}</div>}
    </div>
  );
}
function travelerBrief(o) {
  // one concise sentence for the traveler; the ESP carries full detail
  const t = (o.text || "").replace(/\s+/g, " ").trim();
  const first = t.split(/(?<=[.!?])\s/)[0] || t;
  return first.length > 150 ? first.slice(0, 147) + "…" : first;
}
function IslandTravelerHeader({ p, pn, m, esp, customer }) {
  const cell = { padding: "4px 8px", border: "1px solid #999", fontSize: 10.5 };
  const kc = { ...cell, background: C.gray, fontWeight: 700, fontSize: 9, textTransform: "uppercase" };
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", borderBottom: "2px solid #0B6FB8", paddingBottom: 4, marginBottom: 6 }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <img src={ISLAND_LOGO} alt="Island Components" style={{ height: 30 }} />
          <div><div style={{ fontWeight: 800, fontSize: 15 }}>Job Traveler <span style={{ fontSize: 10, fontWeight: 400, color: "#666" }}>— Shop Copy</span></div>
          <div style={{ fontSize: 8.5, color: "#555" }}>Island Components Group Inc. — A G.W. Lisk Company</div></div></div>
        <div style={{ fontWeight: 800, fontFamily: MONO, fontSize: 14 }}>{pn}</div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        <tr>
          <td style={kc}>Part No.</td><td style={{ ...cell, fontFamily: MONO }}>{pn}</td>
          <td style={kc}>Description</td><td style={cell}>{p.desc || ""}</td>
          <td style={kc}>Rev</td><td style={{ ...cell, fontFamily: MONO }}>{p.rev || "-"}</td>
        </tr>
        <tr>
          <td style={kc}>Job Number</td><td style={cell}><Y on={!m.wo}>{m.wo || "________"}</Y></td>
          <td style={kc}>Customer</td><td style={cell}><Y on={!customer}>{customer || "________"}</Y></td>
          <td style={kc}>Procedure</td><td style={cell}><Y on={esp === "ESP-*"}>{esp}</Y></td>
        </tr>
        <tr>
          <td style={kc}>Date</td><td style={cell}>{m.date}</td>
          <td style={kc}>Qty</td><td style={cell}><Y on>________</Y></td>
          <td style={kc}>Routed By</td><td style={cell}><Y on>________</Y></td>
        </tr>
      </tbody></table>
    </div>
  );
}
function TravelerDocs({ bom, excluded, tops, m, profile, espByPn, customer }) {
  const P = activeProfile(profile);
  const order = buildOrder(bom, excluded, tops);
  const th = { background: C.navy, color: "#fff", padding: "4px 6px", textAlign: "left", fontSize: 9 };
  const td = { border: `1px solid ${C.line}`, padding: "5px 6px", verticalAlign: "top", fontSize: 10 };
  const sig = <div><div style={{ borderBottom: "1px solid #bbb", height: 13, marginBottom: 3 }} /><div style={{ borderBottom: "1px solid #bbb", height: 13 }} /></div>;
  return order.map(pn => {
    const p = bom.parts[pn];
    const { tpl, ops } = opsFor(bom, excluded, pn);
    return (
      <Sheet key={pn}>
        {P.espMode
          ? <IslandTravelerHeader p={p} pn={pn} m={m} esp={espFor(pn, espByPn)} customer={customer} />
          : <DocHeader title={(p.desc || pn).toUpperCase() + " — TRAVELER"} docNo={"TRV-" + pn} pn={pn} m={m} company={P.companyShort} />}
        <H3>GENERAL NOTES</H3>
        <ol style={{ fontSize: 10.5, margin: "0 0 12px 18px" }}>
          {[(P.espMode ? `Job Traveler for ${p.desc || pn} (${pn}). Detailed method per ${espFor(pn, espByPn)}. This traveler is the record of accomplishment, quantities, and sign-off.` : "Sample traveler. Drawing requirements, torque values, cure schedules, test limits, and process specifications require Engineering and Quality approval before production use."),
            "Use only current released drawings, BOMs, specifications, approved supplier parts, calibrated tooling, and in-calibration test equipment.",
            "QA hold points are marked ★ with shaded operation numbers. Do not proceed beyond a hold point without required acceptance and sign-off.",
            "Record all nonconformances in the approved quality system. No unrecorded rework, substitution, or deviation.",
            "Maintain FOD, contamination, ESD, and handling controls appropriate to the hardware and process."].map((n, i) => <li key={i} style={{ margin: "3px 0" }}>{n}</li>)}
        </ol>
        <H3>ROUTING / OPERATIONS</H3>
        <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 14 }}>
          <thead><tr>
            <th style={th}>{P.travelerCols[0]}</th><th style={th}>{P.travelerCols[1]}</th>
            {P.espMode && <th style={th}>{P.travelerCols[2]}</th>}
            <th style={{ ...th, width: P.espMode ? "40%" : "32%" }}>{P.espMode ? P.travelerCols[3] : P.travelerCols[2]}</th>
            {!P.espMode && <th style={{ ...th, width: "22%" }}>{P.travelerCols[3]}</th>}
            <th style={{ ...th, width: 46 }}>{P.travelerCols[4]}</th><th style={{ ...th, width: 46 }}>{P.travelerCols[5]}</th>
            <th style={th}>{P.travelerCols[6]}</th><th style={th}>{P.travelerCols[7]}</th>
          </tr></thead>
          <tbody>
            {ops.map(o => {
              const blank = <div style={{ borderBottom: "1px solid #bbb", height: 13 }} />;
              return (
              <tr key={o.op}>
                <td style={{ ...td, fontFamily: MONO, fontWeight: 700, textAlign: "center", width: 42, background: o.hold ? C.hold : "transparent", color: o.hold ? C.holdInk : "inherit" }}>{o.op}{o.hold ? " ★" : ""}</td>
                <td style={{ ...td, width: 60, fontSize: 9, textTransform: "uppercase" }}>{P.espMode ? islandDept(o.dept) : o.dept}</td>
                {P.espMode && <td style={{ ...td, width: 72, fontSize: 9, fontWeight: 700, textTransform: "uppercase" }}>{P.mapVerb(o.title)}</td>}
                <td style={td}>{P.espMode ? <><b>{o.title}.</b> {travelerBrief(o)} <span style={{ color: "#555" }}>Refer to {espFor(pn, espByPn)} for detailed method.</span></> : o.text}</td>
                {!P.espMode && <td style={td}>{o.accept}</td>}
                <td style={{ ...td, width: 46, verticalAlign: "bottom" }}>{blank}</td>
                <td style={{ ...td, width: 46, verticalAlign: "bottom" }}>{blank}</td>
                <td style={{ ...td, width: 70 }}>{sig}</td>
                <td style={{ ...td, width: 70 }}>{P.espMode ? blank : sig}</td>
              </tr>
            );})}
          </tbody>
        </table>
        {tpl.atp && (<>
          <H3>ACCEPTANCE TEST AND INSPECTION RECORD</H3>
          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 12 }}>
            <thead><tr>
              <th style={{ ...th, width: "32%" }}>Characteristic / Test</th>
              <th style={{ ...th, width: "26%" }}>Requirement</th>
              <th style={th}>Actual Result</th>
              <th style={{ ...th, width: "20%" }}>Inspector / Date</th>
            </tr></thead>
            <tbody>
              {tpl.atp.map((a, i) => (
                <tr key={i}>
                  <td style={td}>{a[0]}</td><td style={td}>{a[1]}</td>
                  <td style={{ ...td, fontFamily: MONO, color: "#bbb" }}>____________</td>
                  <td style={{ ...td, fontFamily: MONO, color: "#bbb" }}>____________</td>
                </tr>
              ))}
            </tbody>
          </table>
        </>)}
        <H3>FINAL DISPOSITION</H3>
        <div style={{ display: "grid", gridTemplateColumns: "150px 1fr 130px 1fr", border: `1px solid ${C.line}`, fontSize: 10.5 }}>
          {[["Accepted for NHA / Stock", "[  ]"], ["Rejected / NCR No.", "____________"],
            ["Final QA Inspector", "____________"], ["Date", "____________"],
            ["Storage Location", "____________"], ["Qty Accepted / Rejected", "______ / ______"]].map(([k, v], i) => (
            <FragmentPair key={i} k={k} v={v} />
          ))}
        </div>
        <Foot company={P.companyShort} />
      </Sheet>
    );
  });
}
function FragmentPair({ k, v }) {
  return (<>
    <div style={{ padding: "5px 8px", border: `1px solid ${C.line}`, background: C.gray, fontWeight: 700 }}>{k}</div>
    <div style={{ padding: "5px 8px", border: `1px solid ${C.line}`, fontFamily: MONO, color: "#bbb" }}>{v}</div>
  </>);
}

/* ---- Work instruction docs ---- */
/* ---- Island ESP process procedure (auto-populated, uncertain fields highlighted) ---- */
function ESPDoc({ bom, excluded, pn, m, esp, customer }) {
  const p = bom.parts[pn];
  const { tpl, ops } = opsFor(bom, excluded, pn);
  const kids = (bom.children[pn] || []).filter(k => !excluded[k.pn]);
  const td = { border: `1px solid ${C.line}`, padding: "4px 8px", fontSize: 10 };
  const th = { background: C.navy, color: "#fff", padding: "4px 6px", textAlign: "left", fontSize: 9 };
  const { tools, docs } = extractToolsAndDocs(ops);
  const drawings = [pn, ...docs.filter(d => !/^ESP/i.test(d))];
  const esps = [...new Set([esp, ...docs.filter(d => /^ESP/i.test(d))])].filter(x => x && x !== "ESP-*");
  const H = ({ n, children }) => <div style={{ fontWeight: 800, fontSize: 12, color: C.navy, margin: "14px 0 6px", borderBottom: `1px solid ${C.line}`, paddingBottom: 2 }}>{n}&nbsp;&nbsp;{children}</div>;
  return (
    <Sheet>
      <IslandBrandBlock />
      {/* cover fields — the three highlighted pull-ins */}
      <div style={{ border: "1px solid #999", padding: "10px 14px", margin: "6px auto 14px", maxWidth: 560, textAlign: "center" }}>
        <div style={{ fontSize: 12, margin: "4px 0" }}><b>Part Number:</b> <span style={{ fontFamily: MONO }}>{pn}</span></div>
        <div style={{ fontSize: 12, margin: "4px 0" }}><b>Customer:</b> <Y on={!customer}>{customer || "________________"}</Y></div>
        <div style={{ fontSize: 15, fontWeight: 800, margin: "8px 0 2px" }}><Y on={esp === "ESP-*"}>{esp}</Y> <span style={{ fontSize: 11, fontWeight: 400 }}>(Procedure Number)</span></div>
        <div style={{ fontSize: 11, marginTop: 4 }}>{(p.desc || "").toUpperCase()}</div>
      </div>
      <div style={{ fontSize: 7.5, color: "#777", textAlign: "center", marginBottom: 10, padding: "0 20px" }}>
        This document contains proprietary information belonging to Island Components Group Inc. and is solely for use by authorized personnel.
      </div>
      <table style={{ width: "auto", borderCollapse: "collapse", margin: "0 0 12px" }}><thead><tr>{["Revision", "Date", "ECO #"].map(h => <th key={h} style={{ ...th, padding: "3px 18px" }}>{h}</th>)}</tr></thead>
        <tbody><tr><td style={{ ...td, fontFamily: MONO }}>{m.rev || "-"}</td><td style={{ ...td, fontFamily: MONO }}>{m.date}</td><td style={{ ...td, fontFamily: MONO }}>{m.eco || ""}</td></tr></tbody></table>

      <H n="1">SCOPE</H>
      <p style={{ fontSize: 11 }}>This procedure provides step-by-step instructions for the fabrication and inspection of <b>{p.desc}</b> (P/N {pn}){customer ? <> for <Y on={custDetGuessFlag(customer, bom)}>{customer}</Y></> : ""}. Operations correspond one-for-one with Job Traveler {pn}.</p>

      <H n="2">APPLICABLE DOCUMENTS</H>
      <table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr><th style={th}>Document</th><th style={th}>Title / Purpose</th></tr></thead><tbody>
        {drawings.map(d => <tr key={d}><td style={{ ...td, fontFamily: MONO }}>{d}</td><td style={td}>{d === pn ? "Assembly drawing" : "Referenced drawing"} <Y on>— verify current revision</Y></td></tr>)}
        {esps.length ? esps.map(d => <tr key={d}><td style={{ ...td, fontFamily: MONO }}>{d}</td><td style={td}>Referenced procedure</td></tr>)
          : <tr><td style={{ ...td, fontFamily: MONO }}><Y on>ESP-*</Y></td><td style={td}><Y on>Related process/test procedure — enter number</Y></td></tr>}
      </tbody></table>

      <H n="3">TOOLS &amp; EQUIPMENT</H>
      {tools.length
        ? <table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr><th style={th}>Tool / Fixture ID</th><th style={th}>Use</th></tr></thead>
            <tbody>{tools.map(t => <tr key={t}><td style={{ ...td, fontFamily: MONO }}>{t}</td><td style={td}><Y on>Referenced in routing — confirm description</Y></td></tr>)}
              <tr><td style={{ ...td }}><Y on>________</Y></td><td style={td}><Y on>Calibrated test equipment as required</Y></td></tr></tbody></table>
        : <p style={{ fontSize: 10.5 }}><Y on>No fixtures were referenced in the routing text — list required tooling, fixtures, and calibrated equipment.</Y></p>}

      <H n="4">MATERIAL</H>
      <table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr>{["Find", "Qty", "Part Number", "Rev", "Description"].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
        <tbody>{kids.map(k => <tr key={k.pn}><td style={{ ...td, textAlign: "center", fontFamily: MONO }}>{k.find}</td><td style={{ ...td, textAlign: "center" }}>{k.qty} {k.uom}</td><td style={{ ...td, fontFamily: MONO }}>{k.pn}</td><td style={{ ...td, textAlign: "center" }}>{k.rev}</td><td style={td}>{k.desc}</td></tr>)}</tbody></table>

      <H n="5">TRAINING REQUIREMENTS</H>
      <p style={{ fontSize: 10.5 }}><Y on>Operators shall be trained and qualified for the operations in this procedure (winding, soldering, impregnation, and inspection as applicable). Enter specific training/certification requirements.</Y></p>

      <H n="6">PROCEDURE</H>
      <p style={{ fontSize: 9.5, color: "#666", fontStyle: "italic", marginBottom: 6 }}>Steps correspond to Job Traveler {pn}. ★ = QA hold point. Record quantities and sign-offs on the traveler.</p>
      {ops.map(o => {
        const vig = vignetteSvg(o.title);
        return (
          <div key={o.op} style={{ marginBottom: 12, breakInside: "avoid" }}>
            <div style={{ display: "grid", gridTemplateColumns: "64px 88px 1fr", background: o.hold ? C.hold : C.navy, color: o.hold ? C.holdInk : "#fff", fontWeight: 700, fontSize: 10.5 }}>
              <div style={{ padding: "5px 8px", fontFamily: MONO }}>OP {o.op}{o.hold ? " ★" : ""}</div>
              <div style={{ padding: "5px 8px" }}>{islandDept(o.dept)} / {activeProfile("island").mapVerb(o.title)}</div>
              <div style={{ padding: "5px 8px" }}>{o.title}{o.hold ? " — QA HOLD" : ""}</div>
            </div>
            {o.sub && o.sub.length ? <ol style={{ margin: "6px 0 6px 22px", fontSize: 10.5 }}>{o.sub.map((x, i) => <li key={i} style={{ margin: "2px 0" }}>{x}</li>)}</ol>
              : <p style={{ margin: "5px 0", fontSize: 10.5 }}>{o.text}</p>}
            <div style={{ border: "1.5px dashed #999", background: "#FBFBF9", padding: "8px 8px 4px", margin: "6px 0" }}>
              <div dangerouslySetInnerHTML={{ __html: vig.svg }} />
              <div style={{ textAlign: "center", fontWeight: 700, fontSize: 9, marginTop: 4 }}>FIGURE {o.op}-1 — {o.title.toUpperCase()}</div>
              <div style={{ textAlign: "center", fontSize: 8.5, color: "#8A6D00", fontWeight: 700 }}>PLACEHOLDER RENDERING — REPLACE WITH PROCESS PHOTO AT RELEASE</div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "110px 1fr", border: `1px solid ${C.line}`, fontSize: 10 }}>
              <div style={{ ...td, background: C.gray, fontWeight: 700 }}>Acceptance</div><div style={td}>{o.accept}</div>
              <div style={{ ...td, background: C.gray, fontWeight: 700 }}>Record</div><div style={td}>{o.record}</div>
            </div>
          </div>
        );
      })}
      <div style={{ textAlign: "center", color: "#999", fontSize: 9, marginTop: 16, letterSpacing: ".04em" }}>UNCONTROLLED WHEN PRINTED · ISLAND COMPONENTS GROUP · {esp}</div>
    </Sheet>
  );
}
function custDetGuessFlag(customer, bom) { const d = detectCustomer(bom); return d.guessed && d.value === customer; }

function WIDocs({ bom, excluded, tops, m, profile, espByPn, customer }) {
  const P = activeProfile(profile);
  const order = buildOrder(bom, excluded, tops);
  const th = { background: C.navy, color: "#fff", padding: "4px 6px", textAlign: "left", fontSize: 9.5 };
  const td = { border: `1px solid ${C.line}`, padding: "3px 6px", verticalAlign: "top", fontSize: 10.5 };
  if (P.espMode) return order.map(pn => <ESPDoc key={pn} bom={bom} excluded={excluded} pn={pn} m={m} esp={espFor(pn, espByPn)} customer={customer} />);
  return order.map(pn => {
    const p = bom.parts[pn];
    const { tpl, ops } = opsFor(bom, excluded, pn);
    const safety = [...new Set(tpl.safety || [])].flatMap(k => SAFETY_TEXT[k] || []);
    const kids = (bom.children[pn] || []).filter(k => !excluded[k.pn]);
    const sec = safety.length ? 4 : 3;
    return (
      <Sheet key={pn}>
        <DocHeader title={"WORK INSTRUCTION — " + (p.desc || pn).toUpperCase()} docNo={"WI-" + pn} pn={pn} extra={"Traveler TRV-" + pn} m={m} company={activeProfile(profile).companyShort} />
        <H3>1&nbsp;&nbsp;PURPOSE AND SCOPE</H3>
        <p style={{ fontSize: 12 }}>Defines the detailed method for building {p.desc}, P/N {pn}, from kitting through final inspection and stocking. Expands traveler TRV-{pn} with step-level detail, tooling, photographs, and workmanship criteria. The traveler remains the record of accomplishment and sign-off; this document is the method.</p>
        <Callout k="NOTE" v={'Sample document. Placeholder requirements ("per drawing," "per released process") shall be replaced with released engineering values before production use.'} />
        <H3>2&nbsp;&nbsp;MATERIALS (PER {pn} BOM)</H3>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>{["Find", "Qty", "Part Number", "Rev", "Description", "Material", "Remarks"].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
          <tbody>
            {kids.map(k => (
              <tr key={k.pn}>
                <td style={{ ...td, textAlign: "center", fontFamily: MONO }}>{k.find}</td>
                <td style={{ ...td, textAlign: "center" }}>{k.qty} {k.uom}</td>
                <td style={{ ...td, fontFamily: MONO }}>{k.pn}</td>
                <td style={{ ...td, textAlign: "center" }}>{k.rev}</td>
                <td style={td}>{k.desc}</td><td style={td}>{k.mat || "—"}</td><td style={td}>{k.rem || ""}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {safety.length > 0 && (<>
          <H3>3&nbsp;&nbsp;SAFETY, HANDLING, AND FOD CONTROLS</H3>
          {safety.map((s, i) => <Callout key={i} k={s[0]} v={s[1]} />)}
        </>)}
        <H3>{sec}&nbsp;&nbsp;DETAILED WORK INSTRUCTIONS</H3>
        <Intro>Operations mirror traveler TRV-{pn}. ★ marks a QA hold point: do not proceed until QA has accepted and signed the traveler.</Intro>
        {ops.map(o => (
          <div key={o.op}>
            <div style={{ display: "grid", gridTemplateColumns: "66px 96px 1fr", border: `1px solid ${C.navy}`, margin: "14px 0 8px" }}>
              <div style={{ padding: "5px 8px", fontWeight: 700, fontSize: 11, fontFamily: MONO, textAlign: "center", background: o.hold ? C.hold : C.navy, color: o.hold ? C.holdInk : "#fff" }}>OP {o.op}{o.hold ? " ★" : ""}</div>
              <div style={{ padding: "5px 8px", fontWeight: 700, fontSize: 9.5, background: o.hold ? C.hold : C.navy, color: o.hold ? C.holdInk : "#fff", display: "flex", alignItems: "center" }}>{o.dept}</div>
              <div style={{ padding: "5px 8px", fontWeight: 700, fontSize: 11, background: o.hold ? C.hold : C.navy, color: o.hold ? C.holdInk : "#fff" }}>{o.title}{o.hold ? " — QA HOLD POINT" : ""}</div>
            </div>
            {(() => { const vig = vignetteSvg(o.title); return (<>
            <div style={{ display: "grid", gridTemplateColumns: "150px 1fr", border: "1px solid #B8860B55", background: "#FFFDF5", fontSize: 10, margin: "6px 0" }}>
              <div style={{ padding: "4px 8px", fontWeight: 700, color: "#8A6D00", borderBottom: "1px solid #B8860B33" }}>DRAWING REFERENCE</div>
              <div style={{ padding: "4px 8px", borderBottom: "1px solid #B8860B33", fontFamily: MONO }}>DWG {pn} — ZONE ______ · VIEW ______</div>
              <div style={{ padding: "4px 8px", fontWeight: 700, color: "#8A6D00", borderBottom: "1px solid #B8860B33" }}>SPEC / PROCEDURE</div>
              <div style={{ padding: "4px 8px", borderBottom: "1px solid #B8860B33", fontFamily: MONO }}>SPEC No. ____________ REV ____ (COMPLETE AT RELEASE)</div>
              <div style={{ padding: "4px 8px", fontWeight: 700, color: "#8A6D00", borderBottom: "1px solid #B8860B33" }}>KEY CHARACTERISTICS</div>
              <div style={{ padding: "4px 8px", borderBottom: "1px solid #B8860B33" }}>{o.accept}</div>
              <div style={{ padding: "4px 8px", fontWeight: 700, color: "#8A6D00" }}>TOOLS / FIXTURES</div>
              <div style={{ padding: "4px 8px" }}>{vig.label} — Tool/Fixture ID ______ (cal. as req'd)</div>
            </div>
            {o.sub && o.sub.length
              ? <ol style={{ margin: "6px 0 8px 22px", fontSize: 11.5 }}>{o.sub.map((x, i) => <li key={i} style={{ margin: "3px 0" }}>{x}</li>)}</ol>
              : <p style={{ margin: "6px 0", fontSize: 11.5 }}>{o.text}</p>}
            <div style={{ border: "1.5px dashed #999", background: "#FBFBF9", padding: "8px 8px 4px", margin: "8px 0" }}>
              <div dangerouslySetInnerHTML={{ __html: vig.svg }} />
              <div style={{ textAlign: "center", fontWeight: 700, fontSize: 9.5, marginTop: 4 }}>FIGURE {o.op}-1 — {o.title.toUpperCase()}</div>
              <div style={{ textAlign: "center", fontSize: 9, color: "#8A6D00", fontWeight: 700 }}>PLACEHOLDER RENDERING — REPLACE WITH PROCESS PHOTO / DRAWING EXTRACT AT RELEASE</div>
              {o.photo && <div style={{ textAlign: "center", fontStyle: "italic", fontSize: 9, color: "#888", marginTop: 2 }}>Intended photo: {o.photo}</div>}
            </div>
            </>); })()}
            {o.callout && <Callout k={o.callout.k} v={o.callout.v} />}
            <div style={{ display: "grid", gridTemplateColumns: "115px 1fr", border: `1px solid ${C.line}`, margin: "8px 0", fontSize: 10.5 }}>
              <div style={{ padding: "4px 8px", border: `1px solid ${C.line}`, background: C.gray, fontWeight: 700 }}>Acceptance</div>
              <div style={{ padding: "4px 8px", border: `1px solid ${C.line}` }}>{o.accept}</div>
              <div style={{ padding: "4px 8px", border: `1px solid ${C.line}`, background: C.gray, fontWeight: 700 }}>Data to Record</div>
              <div style={{ padding: "4px 8px", border: `1px solid ${C.line}` }}>{o.record}</div>
            </div>
          </div>
        ))}
        <H3>{sec + 1}&nbsp;&nbsp;QUALITY RECORDS</H3>
        <p style={{ fontSize: 11.5 }}>Completed traveler TRV-{pn} with all sign-offs; material and adhesive traceability records; cure records with charts where required; test and inspection reports with equipment IDs; nonconformance records per the approved quality system.</p>
        <Foot company={activeProfile(profile).companyShort} />
      </Sheet>
    );
  });
}

/* =====================================================================
   IMPORT PANEL (drag-drop + paste + mapping review)
   ===================================================================== */
function ImportPanel({ onCommit, existingCount, llmCfg }) {
  const [drawPaste, setDrawPaste] = useState(false); const [drawTxt, setDrawTxt] = useState("");
  const [llmBusy, setLlmBusy] = useState(false); const [llmMsg, setLlmMsg] = useState(null);
  const [drag, setDrag] = useState(false);
  const [queue, setQueue] = useState([]); // pending File objects (multi-drop)
  const [batchTotal, setBatchTotal] = useState(0);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [pdfMsg, setPdfMsg] = useState(false);
  const [stage, setStage] = useState(null); // {grid, fileName, sheets, sheetIdx, map, hasHeader, mode, flatParent, merge}
  const fileRef = useRef(null);

  const openGrid = useCallback((grid, fileName, sheets, sheetIdx, wb, ov) => {
    // xls/xlsx exports of ERP BOM reports (Standard sections / kitting-form levels)
    const rep = parseReportGrid(grid);
    if (rep && rep.rows.length) {
      const label = rep.meta.format === "kitting-grid" ? "Kitting form / indentured levels" : `Standard BOM report (${(rep.meta.sections || []).length} sections)`;
      ov = { ...(ov || {}), note: `Detected ${label} · ${rep.rows.length} rows` + (rep.meta.notes.length ? ` · excluded tooling: ${rep.meta.notes.join("; ").slice(0, 120)}` : "") + ((ov && ov.note) ? " · " + ov.note : "") };
      grid = reportToGrid(rep);
    }
    if (!grid.length) return;
    const { map, hasHeader } = guessMapping(grid);
    const mode = (ov && ov.mode) || (map.parent !== undefined ? "parent" : map.level !== undefined ? "level" : "flat");
    setStage(s => ({ grid, fileName, sheets, sheetIdx, wb, map, hasHeader, mode,
      flatParent: (ov && ov.flatParent) || "", rawText: (ov && ov.rawText) || null, note: (ov && ov.note) || null,
      merge: (s && s.merge) || batchActiveRef.current || existingImportedRef.current }));
    setPdfMsg(false);
  }, []);
  const openDrawingText = useCallback((text, fileName) => {
    // ERP BOM reports (Exploded / Standard) take priority over drawing parts-list extraction
    const fmt = detectReportFormat(text);
    if (fmt) {
      const parsed = fmt === "exploded" ? parseExplodedText(text) : parseStandardText(text);
      if (parsed.rows.length) {
        const note = `Detected ${fmt === "exploded" ? "Exploded BOM report (indentured levels)" : `Standard BOM report (${parsed.meta.sections.length} sections)`} · ${parsed.rows.length} rows` +
          (parsed.meta.notes.length ? ` · ${parsed.meta.notes.length} tooling line(s) excluded` : "");
        openGrid(reportToGrid(parsed), fileName || "BOM report", null, 0, null, { rawText: text, note });
        return;
      }
    }
    const ex = extractFromDrawingText(text);
    if (!ex.rows.length) { setPdfMsg("No parts-list rows recognized in that drawing text. Paste the parts-list region (Find/Item, Qty, PN, Description)."); return; }
    const grid = [["Find", "Qty", "UOM", "PartNumber", "Rev", "Description"],
      ...ex.rows.map(r => [r.find, r.qty, r.uom, r.pn, r.rev, r.desc])];
    const note = `Drawing extraction: ${ex.rows.length} rows` + (ex.parent ? ` · title-block PN ${ex.parent}` : " · no title-block PN found — set parent below") +
      (ex.suspect.length ? ` · CHECK OCR-suspect PNs: ${ex.suspect.join(", ")}` : "") + (ex.rejects.length ? ` · ${ex.rejects.length} unparsed line(s) skipped` : "");
    openGrid(grid, fileName || "drawing", null, 0, null, { mode: "flat", flatParent: ex.parent || "", rawText: text, note });
  }, [openGrid]);
  const batchActiveRef = useRef(false);
  const existingImportedRef = useRef(false);

  const handleFiles = useCallback((fileList) => {
    const files = Array.from(fileList);
    if (!files.length) return;
    batchActiveRef.current = files.length > 1;
    setBatchTotal(files.length);
    setQueue(files.slice(1));
    handleFile(files[0]);
  }, []);

  const handleFile = useCallback((file) => {
    const name = file.name.toLowerCase();
    if (/\.(png|jpe?g|bmp|webp)$/.test(name)) {
      setPdfMsg("Running OCR on " + file.name + "…");
      ocrImage(file).then(t => openDrawingText(t, file.name)).catch(e => setPdfMsg("OCR unavailable: " + e.message + " You can still paste the drawing's parts-list text (▦ button below)."));
      advanceQueue(); return;
    }
    if (name.endsWith(".pdf")) {
      setPdfMsg("Extracting " + file.name + "…");
      const rd = new FileReader();
      rd.onload = () => pdfToText(rd.result, msg => setPdfMsg(file.name + ": " + msg + "…"))
        .then(t => openDrawingText(t, file.name))
        .catch(e => { setPdfMsg("PDF extraction unavailable (" + e.message + "). Paste the drawing's parts-list text instead (▦ below), or use the desktop build for fully-local extraction."); setPasteOpen(true); setStage(null); });
      rd.readAsArrayBuffer(file); advanceQueue(); return;
    }
    const reader = new FileReader();
    if (/\.(xlsx|xls|xlsm)$/.test(name)) {
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheets = wb.SheetNames;
          const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheets[0]], { header: 1, defval: "" })
            .map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c));
          openGrid(grid, file.name, sheets, 0, wb);
        } catch (err) { console.error(err); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = e => {
        const { grid } = textToGrid(String(e.target.result));
        openGrid(grid, file.name, null, 0, null);
      };
      reader.readAsText(file);
    }
  }, [openGrid]);

  function switchSheet(idx) {
    const grid = XLSX.utils.sheet_to_json(stage.wb.Sheets[stage.sheets[idx]], { header: 1, defval: "" })
      .map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c));
    const { map, hasHeader } = guessMapping(grid);
    const mode = map.parent !== undefined ? "parent" : map.level !== undefined ? "level" : "flat";
    setStage({ ...stage, grid, sheetIdx: idx, map, hasHeader, mode });
  }

  function advanceQueue() {
    setQueue(q => {
      if (q.length) { setTimeout(() => handleFile(q[0]), 0); return q.slice(1); }
      batchActiveRef.current = false; setBatchTotal(0); return q;
    });
  }
  function commit() {
    const { rows, convIssues } = gridToRows(stage.grid, stage.map, stage.hasHeader, stage.mode, stage.flatParent.trim());
    onCommit(rows, convIssues, stage.merge);
    existingImportedRef.current = true;
    setStage(null); setPasteOpen(false); setPasteText("");
    advanceQueue();
  }
  function skipFile() { setStage(null); advanceQueue(); }

  const nCols = stage ? Math.max(...stage.grid.slice(0, 12).map(r => r.length)) : 0;
  const btn = { border: `1px solid ${C.navy}`, background: "#fff", color: C.navy, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, borderRadius: 2, cursor: "pointer" };
  const btnP = { ...btn, background: C.navy, color: "#fff" };

  return (
    <div>
      {/* drop zone */}
      <div
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }}
        onClick={() => fileRef.current && fileRef.current.click()}
        style={{
          border: `2px dashed ${drag ? C.navy : "#b8b8b2"}`, background: drag ? C.ltblue : "#fff",
          padding: "18px 12px", textAlign: "center", cursor: "pointer", borderRadius: 3, transition: "all .15s",
        }}>
        <div style={{ fontSize: 22, opacity: .5 }}>⤓</div>
        <div style={{ fontWeight: 700, fontSize: 12.5, color: C.navy }}>Drop BOM file(s) or click to browse</div>
        <div style={{ fontSize: 10.5, color: "#888", marginTop: 3 }}>Excel (.xlsx/.xls) · CSV · TSV · TXT — parsed here. Drop several sub-BOM files at once; they queue for review and merge.<br />PDF → use "Paste text" (copy from the PDF viewer).</div>
        <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.xlsm,.csv,.tsv,.txt,.pdf" style={{ display: "none" }}
          onChange={e => { handleFiles(e.target.files); e.target.value = ""; }} />
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
        <button style={btn} onClick={() => { setPasteOpen(!pasteOpen); setPdfMsg(false); }}>{pasteOpen ? "Hide paste box" : "Paste text / CSV"}</button>
        <button style={btn} onClick={() => { setPasteText(SAMPLE_INDENTURED); setPasteOpen(true); setPdfMsg(false); }}>Demo: PDF-style paste</button>
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
        <button style={{ border: `1px solid ${C.line}`, background: "#fff", fontSize: 11, padding: "5px 10px", cursor: "pointer" }} onClick={() => setDrawPaste(v => !v)}>
          ▦ Paste drawing text (deterministic extraction)
        </button>
      </div>
      {drawPaste && (
        <div style={{ marginTop: 8 }}>
          <textarea value={drawTxt} onChange={e => setDrawTxt(e.target.value)} placeholder={"Paste text copied from a vector-PDF drawing (title block + parts list region).\nExample:\nDRAWING NO: ROT-3120\n1  1  SHA-3121  ROTOR SHAFT\n2  8  MAG-3122  PERMANENT MAGNET"}
            style={{ width: "100%", height: 110, fontFamily: MONO, fontSize: 11, border: `1px solid ${C.line}`, padding: 8, boxSizing: "border-box" }} />
          <button style={{ marginTop: 6, border: "none", background: C.navy, color: "#fff", fontSize: 11.5, padding: "6px 14px", cursor: "pointer" }}
            onClick={() => { if (drawTxt.trim()) { openDrawingText(drawTxt, "pasted drawing"); setDrawPaste(false); } }}>Extract parts list</button>
        </div>
      )}
      {pdfMsg && (
        <div style={{ marginTop: 8, background: C.hold, border: `1px solid ${C.holdInk}33`, padding: "8px 10px", fontSize: 11.5, color: C.holdInk }}>
          {typeof pdfMsg === "string" ? pdfMsg : (<><b>PDF detected.</b> This browser preview can't decode PDF binaries — the desktop build will (local pdftotext / pdf.js). For now: open the PDF, select-all, copy, and paste the text below. The indentured-list parser reconstructs the hierarchy from the Level column.</>)}
        </div>
      )}

      {pasteOpen && (
        <div style={{ marginTop: 8 }}>
          <textarea value={pasteText} onChange={e => setPasteText(e.target.value)} spellCheck={false}
            placeholder={"Paste CSV, TSV, or indentured parts-list text copied from a PDF…"}
            style={{ width: "100%", height: 110, fontFamily: MONO, fontSize: 10.5, border: `1px solid ${C.line}`, background: "#fff", padding: 8, resize: "vertical", boxSizing: "border-box" }} />
          <button style={{ ...btnP, marginTop: 6 }} onClick={() => { const { grid } = textToGrid(pasteText); openGrid(grid, "(pasted text)", null, 0, null); }}>
            Parse pasted text ▸
          </button>
        </div>
      )}

      {/* ---- mapping review ---- */}
      {stage && (
        <div style={{ marginTop: 12, background: "#fff", border: `1px solid ${C.navy}`, borderRadius: 3 }}>
          <div style={{ background: C.navy, color: "#fff", padding: "6px 10px", fontSize: 11.5, fontWeight: 700, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            IMPORT REVIEW <span style={{ fontFamily: MONO, fontWeight: 400, opacity: .8 }}>{stage.fileName}</span>
            {batchTotal > 1 && <span style={{ fontSize: 10, opacity: .8, fontWeight: 400 }}>file {batchTotal - queue.length} of {batchTotal}</span>}
            {queue.length > 0 && <button onClick={skipFile} style={{ background: "none", border: "1px solid rgba(255,255,255,.5)", color: "#fff", cursor: "pointer", fontSize: 10, padding: "2px 8px", borderRadius: 2 }}>Skip file</button>}
            <button onClick={() => { setStage(null); setQueue([]); setBatchTotal(0); batchActiveRef.current = false; }} style={{ marginLeft: "auto", background: "none", border: "none", color: "#fff", cursor: "pointer", fontSize: 14 }}>✕</button>
          </div>
          <div style={{ padding: 10 }}>
            {stage.sheets && stage.sheets.length > 1 && (
              <div style={{ marginBottom: 8, fontSize: 11.5 }}>
                Sheet:&nbsp;
                <select value={stage.sheetIdx} onChange={e => switchSheet(+e.target.value)} style={{ fontFamily: MONO, fontSize: 11 }}>
                  {stage.sheets.map((s, i) => <option key={s} value={i}>{s}</option>)}
                </select>
              </div>
            )}

            <div style={{ fontSize: 10.5, color: "#666", marginBottom: 6 }}>
              {stage.grid.length - (stage.hasHeader ? 1 : 0)} data rows · {nCols} columns · header {stage.hasHeader ? "detected" : "not detected"} — verify the column mapping:
            </div>

            {/* mapping selects */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "4px 10px", marginBottom: 8 }}>
              {FIELD_DEFS.map(f => (
                <label key={f.key} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11 }}>
                  <span style={{ width: 82, color: f.key === "pn" ? C.navy : "#666", fontWeight: f.key === "pn" ? 700 : 400 }}>{f.label}</span>
                  <select value={stage.map[f.key] ?? ""} style={{ flex: 1, fontSize: 10.5, fontFamily: MONO }}
                    onChange={e => {
                      const map = { ...stage.map };
                      if (e.target.value === "") delete map[f.key]; else map[f.key] = +e.target.value;
                      setStage({ ...stage, map });
                    }}>
                    <option value="">—</option>
                    {Array.from({ length: nCols }, (_, ci) => (
                      <option key={ci} value={ci}>col {ci + 1}{stage.hasHeader && stage.grid[0][ci] ? ` (${stage.grid[0][ci]})` : ""}</option>
                    ))}
                  </select>
                </label>
              ))}
            </div>

            {/* hierarchy mode */}
            <div style={{ fontSize: 11, marginBottom: 8, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
              <span style={{ fontWeight: 700, color: "#666" }}>Hierarchy:</span>
              {[["parent", "Parent column"], ["level", "Level column"], ["flat", "Single-level"]].map(([v, l]) => (
                <label key={v} style={{ display: "flex", gap: 4, alignItems: "center", cursor: "pointer" }}>
                  <input type="radio" checked={stage.mode === v} onChange={() => setStage({ ...stage, mode: v })} style={{ accentColor: C.navy }} />{l}
                </label>
              ))}
              {stage.note && <div style={{ fontSize: 10.5, color: "#8A6D00", background: "#FFFDF5", border: "1px solid #B8860B44", padding: "5px 8px", marginBottom: 6 }}>{stage.note}</div>}
              {stage.rawText && (
                <div style={{ marginBottom: 6 }}>
                  <button disabled={llmBusy} style={{ border: `1px solid ${C.line}`, background: "#fff", fontSize: 10.5, padding: "4px 10px", cursor: "pointer" }}
                    onClick={async () => {
                      setLlmBusy(true); setLlmMsg(null);
                      try {
                        const csv = await llmNormalize(stage.rawText, llmCfg || { url: "http://localhost:11434", model: "llama3.1" });
                        const g2 = textToGrid(csv);
                        if (g2.grid.length > 1) { openGrid(g2.grid, stage.fileName + " (LLM normalized)", null, 0, null, { mode: "flat", flatParent: stage.flatParent, rawText: stage.rawText, note: "LLM-normalized — verify every part number before commit. Generation itself never uses the LLM." }); }
                        else setLlmMsg("LLM returned no usable rows.");
                      } catch (e) { setLlmMsg("Local LLM unavailable: " + e.message + " (configure endpoint in Settings below)"); }
                      setLlmBusy(false);
                    }}>
                    {llmBusy ? "Normalizing…" : "⟳ LLM normalize (local, optional)"}
                  </button>
                  {llmMsg && <span style={{ fontSize: 10, color: "#B03A00", marginLeft: 8 }}>{llmMsg}</span>}
                </div>
              )}
              {stage.mode === "flat" && (
                <input value={stage.flatParent} onChange={e => setStage({ ...stage, flatParent: e.target.value })}
                  placeholder="Parent PN for all rows" style={{ fontFamily: MONO, fontSize: 10.5, padding: 3, border: `1px solid ${C.line}` }} />
              )}
            </div>

            {/* preview */}
            <div style={{ overflowX: "auto", border: `1px solid ${C.line}`, marginBottom: 8 }}>
              <table style={{ borderCollapse: "collapse", fontSize: 9.5, fontFamily: MONO, width: "100%" }}>
                <thead><tr>
                  {Array.from({ length: nCols }, (_, ci) => {
                    const mapped = FIELD_DEFS.find(f => stage.map[f.key] === ci);
                    return <th key={ci} style={{ background: mapped ? C.ltblue : C.gray, padding: "3px 6px", border: `1px solid ${C.line}`, color: mapped ? C.navy : "#999", whiteSpace: "nowrap" }}>{mapped ? mapped.label : "·"}</th>;
                  })}
                </tr></thead>
                <tbody>
                  {stage.grid.slice(stage.hasHeader ? 1 : 0, (stage.hasHeader ? 1 : 0) + 6).map((r, ri) => (
                    <tr key={ri}>{Array.from({ length: nCols }, (_, ci) => <td key={ci} style={{ padding: "2px 6px", border: `1px solid ${C.line}`, whiteSpace: "nowrap", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis" }}>{r[ci] ?? ""}</td>)}</tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
              <button style={btnP} disabled={stage.map.pn === undefined || (stage.mode === "flat" && !stage.flatParent.trim())} onClick={commit}>
                {stage.map.pn === undefined ? "Map the Part Number column first" : queue.length ? `Commit & next file (${queue.length} queued) ▸` : "Commit import ▸"}
              </button>
              {existingCount > 0 && (
                <label style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 11, cursor: "pointer" }}>
                  <input type="checkbox" checked={stage.merge} onChange={e => setStage({ ...stage, merge: e.target.checked })} style={{ accentColor: C.navy }} />
                  Merge into existing BOM ({existingCount} rows) — for importing sub-BOMs one file at a time
                </label>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* =====================================================================
   MAIN APP
   ===================================================================== */

/* =====================================================================
   STRUCTURE EDITOR — interactive tree correction: drag to re-parent,
   inline edit, add child, delete, mark purchased. Rebuilds BOM live.
   ===================================================================== */
function StructureEditor({ bom, onRows, purchased, onPurchase }) {
  const [editIdx, setEditIdx] = useState(null);
  const [dragIdx, setDragIdx] = useState(null);
  const rows = bom.rows;
  const subtreePNs = pn => {
    const out = new Set([pn]); const walk = q => (bom.children[q] || []).forEach(k => { if (!out.has(k.pn)) { out.add(k.pn); walk(k.pn); } });
    walk(pn); return out;
  };
  const mutate = fn => { const r2 = rows.map(r => ({ ...r })); fn(r2); onRows(r2); setEditIdx(null); };
  const reparent = (srcIdx, targetPn) => {
    const srcPn = rows[srcIdx].pn;
    if (srcPn === targetPn || subtreePNs(srcPn).has(targetPn)) return; // cycle guard
    mutate(r2 => { r2[srcIdx].parent = targetPn; });
  };
  const del = idx => {
    const pn = rows[idx].pn;
    mutate(r2 => {
      r2.splice(idx, 1);
      if (!r2.some(r => r.pn === pn)) { // last instance: cascade orphaned subtree
        const doomed = new Set([pn]); let changed = true;
        while (changed) { changed = false; r2.forEach(r => { if (doomed.has(r.parent) && !doomed.has(r.pn)) { doomed.add(r.pn); changed = true; } }); }
        for (let i = r2.length - 1; i >= 0; i--) if (doomed.has(r2[i].parent)) r2.splice(i, 1);
      }
    });
  };
  const addChild = pn => mutate(r2 => {
    const sibs = r2.filter(r => r.parent === pn).length;
    r2.push({ parent: pn, find: String((sibs + 1) * 10), qty: "1", uom: "EA", pn: "NEW-" + String(100 + Math.floor(Math.random() * 900)), rev: "-", desc: "New Part — edit me", mb: "", mat: "", rem: "" });
  });
  const F = ["pn", "desc", "qty", "uom", "rev", "mb"];
  const nodeRow = (pn, row, idx, depth) => {
    const p = bom.parts[pn] || { pn, desc: "", rev: "-" };
    const kids = bom.children[pn] || [];
    const missing = !kids.length && isAssemblyLike(p) && !purchased[pn];
    const isPurch = !kids.length && isAssemblyLike(p) && purchased[pn];
    const editing = editIdx === idx && idx !== null;
    return (
      <div key={pn + "_" + idx + "_" + depth}>
        <div draggable={idx !== null} onDragStart={() => setDragIdx(idx)} onDragOver={e => e.preventDefault()}
          onDrop={e => { e.preventDefault(); if (dragIdx !== null && dragIdx !== idx) reparent(dragIdx, pn); setDragIdx(null); }}
          style={{ display: "flex", alignItems: "center", gap: 6, padding: "3px 6px", marginLeft: depth * 18, borderLeft: depth ? `2px solid ${C.line}` : "none",
            background: missing ? "#FFF9E8" : editing ? "#EAF0FA" : "transparent", fontSize: 11.5, cursor: idx !== null ? "grab" : "default" }}>
          <span style={{ color: "#bbb", fontSize: 10 }}>{idx !== null ? "⠿" : "▣"}</span>
          {!editing ? (<>
            <b style={{ fontFamily: MONO, color: kids.length ? C.navy2 : "#222" }}>{pn}</b>
            <span style={{ color: "#555", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 240 }}>{p.desc}</span>
            {row && <span style={{ fontFamily: MONO, fontSize: 10, color: "#777" }}>×{row.qty} {row.uom}</span>}
            {missing && <span style={{ fontSize: 9.5, color: "#B8860B", fontWeight: 700 }}>▲ NO BOM</span>}
            {isPurch && <span style={{ fontSize: 9.5, color: "#666" }}>(PURCHASED)</span>}
            <span style={{ flex: 1 }} />
            {idx !== null && <button title="Edit" onClick={() => setEditIdx(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11 }}>✎</button>}
            <button title="Add child" onClick={() => addChild(pn)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 12, color: C.navy2 }}>＋</button>
            {(!kids.length && isAssemblyLike(p)) && <button title="Toggle purchased" onClick={() => onPurchase(pn)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 10 }}>🛒</button>}
            {idx !== null && <button title="Delete" onClick={() => del(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11, color: "#B03A00" }}>🗑</button>}
          </>) : (<>
            {F.map(f => (
              <input key={f} defaultValue={f === "pn" ? row.pn : f === "desc" ? (bom.parts[row.pn] || {}).desc || "" : row[f] || ""}
                onChange={e => { row["_" + f] = e.target.value; }} placeholder={f}
                style={{ fontFamily: f === "pn" ? MONO : "inherit", fontSize: 10.5, border: `1px solid ${C.line}`, padding: "2px 4px", width: f === "desc" ? 170 : f === "pn" ? 90 : 40 }} />
            ))}
            <button onClick={() => mutate(r2 => { const r = r2[idx];
              F.forEach(f => { if (row["_" + f] !== undefined) { if (f === "pn") { const old = r.pn; r.pn = row._pn.toUpperCase(); r2.forEach(x => { if (x.parent === old) x.parent = r.pn; }); } else r[f] = row["_" + f]; } });
            })} style={{ border: "none", background: C.navy, color: "#fff", fontSize: 10, padding: "3px 8px", cursor: "pointer" }}>✓</button>
            <button onClick={() => setEditIdx(null)} style={{ border: `1px solid ${C.line}`, background: "#fff", fontSize: 10, padding: "3px 6px", cursor: "pointer" }}>✕</button>
          </>)}
        </div>
        {kids.map(k => { const i2 = rows.findIndex(r => r === k || (r.pn === k.pn && r.parent === pn && r.find === k.find)); return nodeRow(k.pn, k, i2 >= 0 ? i2 : null, depth + 1); })}
      </div>
    );
  };
  return (
    <div style={{ border: `1px solid ${C.line}`, background: "#fff", padding: 8, maxHeight: 340, overflowY: "auto" }}>
      <div style={{ fontSize: 10, color: "#888", marginBottom: 6 }}>Drag ⠿ onto a new parent to re-parent · ✎ edit fields · ＋ add child · 🛒 toggle purchased · 🗑 delete (cascades if last instance). Cycles are blocked. Changes re-validate live.</div>
      {bom.tops.map(t => nodeRow(t, null, null, 0))}
    </div>
  );
}

/* =====================================================================
   OFFICE EXPORT + EDIT MODE
   - Word export: pane HTML -> Word-compatible .doc (SVGs converted to PNG)
   - Excel export: parts-list table -> .xlsx via SheetJS
   - Optional export folder via File System Access API (Chrome/Edge);
     falls back to normal downloads elsewhere.
   ===================================================================== */
async function ensureDocx() {
  if (typeof window === "undefined") throw new Error("No browser environment.");
  if (window.docx) return window.docx;
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/docx.js");
  if (!window.docx) throw new Error("docx library did not initialize.");
  return window.docx;
}
/* rasterize an SVG element to a PNG data-url + pixel dims (for embedding in docx) */
async function svgToPngData(svg) {
  const xml = new XMLSerializer().serializeToString(svg);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
  const img = new Image();
  await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = url; });
  const r = svg.getBoundingClientRect();
  const scale = 2, w = Math.max(60, Math.round(r.width)), h = Math.max(30, Math.round(r.height));
  const cv = document.createElement("canvas"); cv.width = w * scale; cv.height = h * scale;
  const ctx = cv.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);
  ctx.drawImage(img, 0, 0, cv.width, cv.height);
  return { dataUrl: cv.toDataURL("image/png"), w, h };
}
function dataUrlToUint8(dataUrl) {
  const b64 = dataUrl.split(",")[1]; const bin = atob(b64);
  const arr = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}
const rgb = c => { const m = (c || "").match(/rg.*?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/); if (!m) return null; return ((+m[1]) << 16 | (+m[2]) << 8 | (+m[3])).toString(16).padStart(6, "0"); };
/* Convert the rendered document DOM into a real .docx (tables preserved, drawings embedded) */
async function domToDocx(rootEl, title, landscape) {
  const D = await ensureDocx();
  // pre-rasterize all svgs
  const svgMap = new Map();
  const svgs = [...rootEl.querySelectorAll("svg")];
  for (const s of svgs) { try { svgMap.set(s, await svgToPngData(s)); } catch (e) {} }
  const children = [];
  const runsFromEl = el => {
    // collect text with basic bold/color from inline styles
    const runs = [];
    const walk = (node, inherited) => {
      if (node.nodeType === 3) { const t = node.textContent.replace(/\s+/g, " "); if (t.trim()) runs.push(new D.TextRun({ text: t, bold: inherited.bold, color: inherited.color, size: inherited.size, highlight: inherited.hl })); return; }
      if (node.nodeType !== 1) return;
      const st = node.getAttribute && node.getAttribute("style") || "";
      const cs = window.getComputedStyle(node);
      const bold = inherited.bold || /font-weight:\s*(bold|[6-9]00)/.test(st) || +cs.fontWeight >= 600;
      const col = rgb(cs.color) || inherited.color;
      const hl = /background(-color)?:\s*#?FFF200|background(-color)?:\s*rgb\(255,\s*242,\s*0/i.test(st) ? "yellow" : inherited.hl;
      const next = { bold, color: col, size: inherited.size, hl };
      node.childNodes.forEach(c => walk(c, next));
    };
    el.childNodes.forEach(c => walk(c, { bold: false, color: undefined, size: 18, hl: undefined }));
    return runs;
  };
  const tableToDocx = tbl => {
    const rows = [...tbl.rows].map(tr => {
      const cells = [...tr.cells].map(tc => {
        const cs = window.getComputedStyle(tc);
        const shade = rgb(cs.backgroundColor);
        return new D.TableCell({
          shading: shade && shade !== "ffffff" ? { fill: shade } : undefined,
          margins: { top: 40, bottom: 40, left: 60, right: 60 },
          children: [new D.Paragraph({ children: runsFromEl(tc).length ? runsFromEl(tc) : [new D.TextRun("")] })],
        });
      });
      return new D.TableRow({ children: cells });
    });
    return new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows });
  };
  // walk top-level blocks of each "Sheet" div in order
  const sheets = rootEl.children.length ? [...rootEl.children] : [rootEl];
  sheets.forEach((sheet, si) => {
    if (si > 0) children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    const walkBlock = el => {
      for (const node of el.children) {
        const tag = node.tagName;
        if (tag === "TABLE") { children.push(tableToDocx(node)); children.push(new D.Paragraph({ text: "" })); continue; }
        const svg = node.tagName === "SVG" ? node : node.querySelector && node.querySelector("svg");
        if (svg && svgMap.has(svg)) {
          const im = svgMap.get(svg);
          const maxW = landscape ? 900 : 640; const w = Math.min(maxW, im.w); const h = im.h * (w / im.w);
          children.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.ImageRun({ data: dataUrlToUint8(im.dataUrl), transformation: { width: w, height: h } })] }));
          continue;
        }
        // CSS grid with a fixed column count -> render as a Word table so it doesn't flatten
        const cs0 = window.getComputedStyle(node);
        if (cs0.display === "grid" && node.children.length > 1) {
          const tmpl = cs0.gridTemplateColumns || "";
          const ncol = tmpl.split(" ").filter(Boolean).length;
          if (ncol >= 2 && node.children.length >= ncol) {
            const cells = [...node.children];
            const rows = [];
            for (let i = 0; i < cells.length; i += ncol) {
              const rowCells = cells.slice(i, i + ncol).map(c => {
                const ccs = window.getComputedStyle(c); const shade = rgb(ccs.backgroundColor);
                return new D.TableCell({ shading: shade && shade !== "ffffff" ? { fill: shade } : undefined, margins: { top: 30, bottom: 30, left: 50, right: 50 }, children: [new D.Paragraph({ children: runsFromEl(c).length ? runsFromEl(c) : [new D.TextRun("")] })] });
              });
              while (rowCells.length < ncol) rowCells.push(new D.TableCell({ children: [new D.Paragraph("")] }));
              rows.push(new D.TableRow({ children: rowCells }));
            }
            children.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows }));
            children.push(new D.Paragraph({ text: "" }));
            continue;
          }
        }
        if (node.querySelector && (node.querySelector("table") || node.querySelector("svg"))) { walkBlock(node); continue; }
        const runs = runsFromEl(node);
        if (runs.length) {
          const cs = window.getComputedStyle(node);
          const big = parseFloat(cs.fontSize) >= 15 || +cs.fontWeight >= 700;
          children.push(new D.Paragraph({ spacing: { after: 60 }, children: runs, heading: big && node.textContent.length < 70 ? D.HeadingLevel.HEADING_3 : undefined }));
        }
      }
    };
    walkBlock(sheet);
  });
  const doc = new D.Document({
    sections: [{
      properties: { page: { size: landscape ? { orientation: D.PageOrientation.LANDSCAPE, width: 15840, height: 12240 } : { width: 12240, height: 15840 }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
      children,
    }],
  });
  return await D.Packer.toBlob(doc);
}
let EXPORT_DIR = null; // FileSystemDirectoryHandle when user picks a folder
async function pickExportFolder() {
  if (!window.showDirectoryPicker) throw new Error("Folder access needs Chrome or Edge. Exports will download normally instead.");
  EXPORT_DIR = await window.showDirectoryPicker({ mode: "readwrite" });
  return EXPORT_DIR.name;
}
async function saveOut(filename, blob) {
  if (EXPORT_DIR) {
    try {
      const fh = await EXPORT_DIR.getFileHandle(filename, { create: true });
      const ws = await fh.createWritable(); await ws.write(blob); await ws.close();
      return "folder";
    } catch (e) { /* fall through to download */ }
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return "download";
}
async function exportPaneAsPDF(paneEl, title, pageSize) {
  // Print the exact preview via the browser's own renderer -> user picks "Save as PDF".
  // pageSize: "letter" | "tabloid-landscape" | "tabloid-portrait"
  const svgClones = [];
  // inline computed background so print shows shading
  const win = window.open("", "_blank");
  if (!win) throw new Error("Popup blocked — allow popups for this site to export PDF.");
  const sizeCSS = pageSize === "tabloid-landscape" ? "17in 11in"
    : pageSize === "tabloid-portrait" ? "11in 17in" : "8.5in 11in";
  const html = paneEl.innerHTML;
  win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>
<style>
  @page { size: ${sizeCSS}; margin: 0.4in; }
  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; background: #fff; }
  svg { max-width: none !important; }
  @media print { .noprint { display: none !important; } }
  .barp { position: fixed; top: 0; left: 0; right: 0; background: #1F3864; color: #fff; padding: 10px 16px; font: 13px Segoe UI, Arial; z-index: 99; display: flex; gap: 12px; align-items: center; }
  .barp button { background: #fff; color: #1F3864; border: none; padding: 6px 14px; font-weight: 700; border-radius: 3px; cursor: pointer; }
  .content { margin-top: 0; }
  @media screen { .content { margin-top: 52px; padding: 16px; } }
</style></head><body>
<div class="barp noprint"><b>DocWorks — Print to PDF</b><button onclick="window.print()">🖨 Print / Save as PDF</button><span style="font-weight:400;font-size:12px">In the dialog choose "Save as PDF" as the destination${pageSize !== "letter" ? ' and set paper to <b>Tabloid / 11×17</b> ' + (pageSize.includes("landscape") ? "Landscape" : "Portrait") : ""}.</span></div>
<div class="content">${html}</div>
</body></html>`);
  win.document.close();
  // give layout a tick, then auto-open print dialog
  setTimeout(() => { try { win.focus(); win.print(); } catch (e) {} }, 600);
  return "print";
}
async function exportPaneAsWord(paneEl, filename, landscape) {
  try {
    const blob = await domToDocx(paneEl, filename, landscape);
    return saveOut(filename.endsWith(".docx") ? filename : filename + ".docx", blob);
  } catch (e) {
    // offline / CDN blocked: fall back to a Word-openable HTML with inline table styles
    const clone = paneEl.cloneNode(true);
    clone.querySelectorAll("svg").forEach(s => { const note = document.createElement("div"); note.textContent = "[ drawing — view in app or PDF ]"; note.style.cssText = "border:1px dashed #999;padding:10px;text-align:center;color:#888"; s.replaceWith(note); });
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><style>@page{size:${landscape ? "17in 11in" : "8.5in 11in"};margin:.6in}body{font-family:Segoe UI,Calibri,Arial;font-size:10pt}table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:3pt 5pt;font-size:9pt;vertical-align:top}</style></head><body>${clone.innerHTML}</body></html>`;
    const where = await saveOut(filename.replace(/\.docx$/, "") + ".doc", new Blob(["\ufeff" + html], { type: "application/msword" }));
    return where;
  }
}
function exportTableAsXlsx(paneEl, filename) {
  const tables = paneEl.querySelectorAll("table");
  if (!tables.length) throw new Error("No table found to export.");
  const wb = XLSX.utils.book_new();
  tables.forEach((t, i) => {
    const ws = XLSX.utils.table_to_sheet(t);
    XLSX.utils.book_append_sheet(wb, ws, ("Sheet" + (i + 1)).slice(0, 31));
  });
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return saveOut(filename.endsWith(".xlsx") ? filename : filename + ".xlsx", new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
}

function DocWorks() {
  const [bom, setBom] = useState(() => parseCsvBOM(SAMPLE));
  const [srcLabel, setSrcLabel] = useState("EZ Motors sample");
  const [activeCfg, setActiveCfg] = useState("actuator");
  const [excluded, setExcluded] = useState({});
  const [manualTop, setManualTop] = useState("");
  const [wo, setWo] = useState(""); const [sn, setSn] = useState("");
  const [rev, setRev] = useState("1"); const [eco, setEco] = useState("ECO-0001"); const [change, setChange] = useState("INITIAL RELEASE (GENERATED)");
  const [showEditor, setShowEditor] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [check, setCheck] = useState(null);
  const [customDecls, setCustomDecls] = useState([]);
  const [llmUrl, setLlmUrl] = useState("http://localhost:11434"); const [llmModel, setLlmModel] = useState("llama3.1");
  const [llmStatus, setLlmStatus] = useState(null);
  const projFileRef = useRef(null); const tplFileRef = useRef(null);
  const applyRows = rows => { setBom(buildBOM(rows, [])); setGenerated(null); setCheck(null); setSrcLabel(s => /\(edited\)$/.test(s) ? s : s + " (edited)"); };
  const download = (name, obj) => { const b = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" }); const a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 500); };
  const exportProject = () => download("docworks_project.json", { version: "0.9", rows: bom.rows, srcLabel, activeCfg, excluded, purchased, meta: { wo, sn, prog, date, rev, eco, change, profile, espByPn, customer, sheetSize }, customTemplates: customDecls });
  const importProject = f => { const rd = new FileReader(); rd.onload = () => { try {
      const j = JSON.parse(rd.result);
      if (!Array.isArray(j.rows)) throw new Error("no rows[]");
      setCustomDecls(j.customTemplates || []); setCustomTemplates(j.customTemplates || []);
      setBom(buildBOM(j.rows, [])); setSrcLabel(j.srcLabel || f.name); setActiveCfg(j.activeCfg || "actuator");
      setExcluded(j.excluded || {}); setPurchased(j.purchased || {});
      const mm = j.meta || {}; if (mm.profile) setProfile(mm.profile); if (mm.sheetSize) setSheetSize(mm.sheetSize); if (mm.espByPn) setEspByPn(mm.espByPn); if (mm.customer) setCustomerOverride(mm.customer); setWo(mm.wo || ""); setSn(mm.sn || ""); setProg(mm.prog || "Sample Program"); setDate(mm.date || date); setRev(mm.rev || "1"); setEco(mm.eco || "ECO-0001"); setChange(mm.change || "INITIAL RELEASE (GENERATED)");
      setGenerated(null); setCheck(null);
    } catch (e) { alert("Project import failed: " + e.message); } }; rd.readAsText(f); };
  const importTemplates = f => { const rd = new FileReader(); rd.onload = () => { try {
      const j = JSON.parse(rd.result); const arr = Array.isArray(j) ? j : j.templates;
      if (!Array.isArray(arr)) throw new Error("expected an array of template declarations");
      arr.forEach(d => { if (!d.id || !d.match || !Array.isArray(d.ops)) throw new Error("each template needs id, match, ops[]"); });
      setCustomDecls(arr); setCustomTemplates(arr); setGenerated(null); setCheck(null);
    } catch (e) { alert("Template import failed: " + e.message); } }; rd.readAsText(f); };
  const [prog, setProg] = useState("Sample Program");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [tab, setTab] = useState("tree");
  const [generated, setGenerated] = useState(null);
  const [purchased, setPurchased] = useState({}); // pn -> true (user marked "purchased, no BOM expected")
  const [profile, setProfile] = useState("island"); // output profile: ez | island
  const [sheetSize, setSheetSize] = useState("tabloid"); // letter (multi-sheet) | tabloid (11x17 one sheet)
  const [editMode, setEditMode] = useState(false);
  const [exportDirName, setExportDirName] = useState("");
  const [exportMsg, setExportMsg] = useState("");
  const docsRef = useRef(null);
  // table row editing in edit mode: right-click a row -> insert/duplicate/delete
  useEffect(() => {
    if (!editMode) return;
    const root = docsRef.current; if (!root) return;
    const onCtx = e => {
      const tr = e.target.closest && e.target.closest("tr");
      if (!tr || !root.contains(tr)) return;
      e.preventDefault();
      const tbody = tr.parentNode;
      const menu = document.createElement("div");
      menu.style.cssText = `position:fixed;left:${e.clientX}px;top:${e.clientY}px;z-index:9999;background:#fff;border:1px solid #999;box-shadow:0 3px 12px rgba(0,0,0,.2);font:12px Segoe UI,Arial;border-radius:3px;overflow:hidden`;
      const mk = (label, fn) => { const b = document.createElement("div"); b.textContent = label; b.style.cssText = "padding:7px 14px;cursor:pointer;white-space:nowrap"; b.onmouseenter = () => b.style.background = "#EEF"; b.onmouseleave = () => b.style.background = "#fff"; b.onclick = () => { fn(); document.body.removeChild(menu); }; menu.appendChild(b); };
      mk("Insert row above", () => { const c = tr.cloneNode(true); c.querySelectorAll("td,th").forEach(td => { if (!/^(Op|Step|\d+)/.test(td.textContent)) td.textContent = ""; }); tbody.insertBefore(c, tr); });
      mk("Insert row below", () => { const c = tr.cloneNode(true); c.querySelectorAll("td,th").forEach(td => { td.textContent = ""; }); tbody.insertBefore(c, tr.nextSibling); });
      mk("Duplicate row", () => { tbody.insertBefore(tr.cloneNode(true), tr.nextSibling); });
      mk("Delete row", () => { tbody.removeChild(tr); });
      document.body.appendChild(menu);
      const close = ev => { if (!menu.contains(ev.target)) { if (menu.parentNode) document.body.removeChild(menu); document.removeEventListener("mousedown", close); } };
      setTimeout(() => document.addEventListener("mousedown", close), 0);
    };
    root.addEventListener("contextmenu", onCtx);
    return () => root.removeEventListener("contextmenu", onCtx);
  }, [editMode]);
  const [espByPn, setEspByPn] = useState({});    // Island: per-assembly ESP number overrides
  const custDet = useMemo(() => detectCustomer(bom), [bom]);
  const [customerOverride, setCustomerOverride] = useState("");
  const customer = customerOverride || custDet.value || "";

  const configs = useMemo(() => makeConfigs(bom), [bom]);
  const cfg = configs.find(c => c.id === activeCfg);
  const cfgTops = bom ? cfg.tops() : null;
  const needManual = bom && (!cfgTops || !cfgTops.length);

  const options = useMemo(() => {
    if (!bom || !cfg.hasOptions || !cfgTops || !cfgTops.length) return [];
    const kids = bom.children[cfgTops[0]] || [];
    const found = [];
    OPTION_CLASSES.forEach(oc => {
      kids.filter(k => oc.re.test((bom.parts[k.pn] || {}).desc || "")).forEach(k =>
        found.push({ ...oc, pn: k.pn, desc: bom.parts[k.pn].desc }));
    });
    return found;
  }, [bom, activeCfg, cfgTops]);

  const asmList = useMemo(() => bom ? Object.values(bom.parts).filter(p => bom.children[p.pn] && bom.children[p.pn].length) : [], [bom]);

  // leaf parts that look like assemblies but have no BOM loaded under them
  const missingBoms = useMemo(() => {
    if (!bom) return [];
    return Object.values(bom.parts).filter(p =>
      !(bom.children[p.pn] && bom.children[p.pn].length) && /assembl|assy|\bkit\b/i.test(p.desc || ""));
  }, [bom]);
  const unresolvedMissing = missingBoms.filter(p => !purchased[p.pn]);

  function handleImport(rows, convIssues, merge) {
    const allRows = merge && bom ? [...bom.rows, ...rows] : rows;
    setBom(buildBOM(allRows, convIssues));
    setSrcLabel(merge ? srcLabel + " + import" : "imported");
    setGenerated(null); setExcluded({}); setManualTop("");
    if (!merge) setPurchased({});
  }
  function loadSample() {
    setBom(parseCsvBOM(SAMPLE)); setSrcLabel("EZ Motors sample");
    setGenerated(null); setExcluded({}); setManualTop(""); setPurchased({});
  }
  function doGenerate() {
    const tops = (cfgTops && cfgTops.length) ? cfgTops : (manualTop ? [manualTop] : []);
    if (!tops.length) return;
    setCheck(runDocumentCheck(bom, excluded, tops, purchased));
    setGenerated({ tops, cfgName: cfg.name, excluded: { ...excluded }, purchased: { ...purchased } });
    setTab("tree");
  }

  const m = { wo: wo || "____________", sn: sn || "____________", prog, date, rev: rev || "1", eco: eco || "ECO-0001", change: change || "INITIAL RELEASE (GENERATED)" };
  const btn = { border: `1px solid ${C.navy}`, background: "#fff", color: C.navy, padding: "6px 12px", fontSize: 12, fontWeight: 600, borderRadius: 2, cursor: "pointer" };
  const btnP = { ...btn, background: C.navy, color: "#fff" };
  const stepNum = { width: 24, height: 24, borderRadius: 3, background: C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: MONO, fontSize: 12, fontWeight: 700, flexShrink: 0 };
  const stepH = { fontSize: 12, letterSpacing: ".08em", textTransform: "uppercase", color: C.navy, fontWeight: 700 };
  const inputS = { width: "100%", border: `1px solid ${C.line}`, padding: 5, fontSize: 12, fontFamily: MONO, background: "#fff", marginTop: 3, boxSizing: "border-box" };
  const lblS = { fontSize: 10, textTransform: "uppercase", letterSpacing: ".05em", color: "#888", display: "block" };

  return (
    <div style={{ fontFamily: '"Segoe UI",system-ui,sans-serif', background: C.backdrop, color: C.ink, fontSize: 14, minHeight: "100vh", display: "flex", flexDirection: "column" }}>
      <div style={{ background: C.navy, color: "#fff", padding: "10px 18px", display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ fontWeight: 800, letterSpacing: ".12em", fontSize: 16 }}>DOC<span style={{ color: "#F2C14E" }}>WORKS</span></div>
        <div style={{ opacity: .75, fontSize: 11.5, borderLeft: "1px solid rgba(255,255,255,.3)", paddingLeft: 14 }}>BOM / drawing import → Family Tree · Parts List · Traveler · Work Instruction | 100% local</div>
        <div style={{ marginLeft: "auto", fontSize: 10.5, opacity: .65, fontFamily: MONO }}>v0.17 PROTOTYPE</div>
      </div>

      <div style={{ display: "flex", flex: 1, minHeight: 0, flexWrap: "wrap" }}>
        {/* LEFT RAIL */}
        <div style={{ width: 400, minWidth: 310, flexShrink: 0, background: C.paper, borderRight: `1px solid ${C.line}`, padding: 16, overflowY: "auto", maxHeight: "calc(100vh - 46px)" }}>

          {/* Step 1 — import */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={stepNum}>1</div><div style={stepH}>Bill of Materials</div>
            </div>
            <div style={{ borderLeft: `2px solid ${C.line}`, marginLeft: 11, paddingLeft: 20 }}>
              <ImportPanel onCommit={handleImport} existingCount={bom ? bom.rows.length : 0} llmCfg={{ url: llmUrl, model: llmModel }} />
              <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                <button style={btn} onClick={loadSample}>Reset to sample</button>
              </div>
              {bom && (
                <div style={{ marginTop: 10, fontSize: 11.5 }}>
                  {bom.issues.length
                    ? <div style={{ color: C.warnInk }}>Loaded with {bom.issues.length} validation finding(s):
                        <ul style={{ margin: "5px 0 0 16px", padding: 0 }}>{bom.issues.slice(0, 8).map((x, i) => <li key={i}>{x}</li>)}{bom.issues.length > 8 && <li>…{bom.issues.length - 8} more</li>}</ul></div>
                    : <div style={{ color: C.stamp, fontWeight: 600 }}>✓ BOM loaded clean — no validation findings.</div>}
                  <div style={{ fontFamily: MONO, fontSize: 10.5, color: "#666", marginTop: 5 }}>
                    source: {srcLabel} · {Object.keys(bom.parts).length} parts · {asmList.length} assemblies · top: {bom.tops.join(", ") || "(none)"}
                  </div>
                </div>
              )}
              {missingBoms.length > 0 && (
                <div style={{ marginTop: 10, background: "#FFF9E8", border: "1px solid #B8860B55", borderRadius: 3, padding: "8px 10px" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#8A6D00", marginBottom: 5 }}>
                    ▲ {missingBoms.length} assembl{missingBoms.length > 1 ? "ies" : "y"} named with no BOM loaded
                  </div>
                  <div style={{ fontSize: 10.5, color: "#8A6D00", marginBottom: 6 }}>
                    Drop the missing sub-BOM file(s) above to merge them in, or mark as purchased (bought complete — no internal BOM expected).
                  </div>
                  {missingBoms.map(p => (
                    <div key={p.pn} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11, padding: "3px 0", borderTop: "1px dashed #B8860B33" }}>
                      <span style={{ fontFamily: MONO, fontWeight: 700 }}>{p.pn}</span>
                      <span style={{ color: "#666", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{p.desc}</span>
                      {purchased[p.pn]
                        ? <button onClick={() => { const x = { ...purchased }; delete x[p.pn]; setPurchased(x); setGenerated(null); }}
                            style={{ border: "1px solid #999", background: "#eee", color: "#555", fontSize: 9.5, padding: "2px 7px", cursor: "pointer", borderRadius: 2 }}>PURCHASED ✓ (undo)</button>
                        : <button onClick={() => { setPurchased({ ...purchased, [p.pn]: true }); setGenerated(null); }}
                            style={{ border: "1px solid #B8860B", background: "#fff", color: "#8A6D00", fontSize: 9.5, padding: "2px 7px", cursor: "pointer", borderRadius: 2 }}>Mark purchased</button>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Step 2 — configuration */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={{ marginBottom: 14 }}>
                <button style={{ ...btn, width: "100%", fontSize: 11 }} onClick={() => setShowEditor(v => !v)}>
                  {showEditor ? "▾ Hide" : "▸ Show"} Structure Editor (drag-drop correction)
                </button>
                {showEditor && <div style={{ marginTop: 8 }}>
                  <StructureEditor bom={bom} onRows={applyRows} purchased={purchased} onPurchase={pn => setPurchased(pu => ({ ...pu, [pn]: !pu[pn] }))} />
                </div>}
              </div>
              <div style={stepNum}>2</div><div style={stepH}>Build Configuration</div>
            </div>
            <div style={{ borderLeft: `2px solid ${C.line}`, marginLeft: 11, paddingLeft: 20 }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {configs.map(c => {
                  const ok = !bom || c.tops();
                  const sel = c.id === activeCfg;
                  return (
                    <label key={c.id} style={{
                      display: "flex", gap: 10, alignItems: "flex-start", border: `1px solid ${sel ? C.navy : C.line}`,
                      boxShadow: sel ? `inset 3px 0 0 ${C.navy}` : "none", background: "#fff", padding: "8px 10px", borderRadius: 2,
                      cursor: "pointer", opacity: ok ? 1 : .55,
                    }}>
                      <input type="radio" name="cfg" checked={sel} onChange={() => { setActiveCfg(c.id); setExcluded({}); setGenerated(null); }} style={{ marginTop: 3, accentColor: C.navy }} />
                      <div>
                        <div style={{ fontWeight: 700, fontSize: 12.5 }}>{c.name}{!ok && <span style={{ color: "#bbb", fontWeight: 400 }}> (no match in BOM)</span>}</div>
                        <div style={{ fontSize: 11, color: "#666", marginTop: 2 }}>{c.desc}</div>
                      </div>
                    </label>
                  );
                })}
              </div>
              {options.length > 0 && (
                <div style={{ marginTop: 10, background: "#fff", border: `1px dashed ${C.line}`, padding: 10 }}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: ".08em", color: "#888", marginBottom: 6, fontWeight: 700 }}>Included options</div>
                  {options.map(f => (
                    <label key={f.pn} style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, padding: "3px 0", cursor: "pointer" }}>
                      <input type="checkbox" checked={!excluded[f.pn]} style={{ accentColor: C.navy }}
                        onChange={e => { const x = { ...excluded }; if (e.target.checked) delete x[f.pn]; else x[f.pn] = true; setExcluded(x); setGenerated(null); }} />
                      {f.label} — {f.desc}
                      <span style={{ fontFamily: MONO, fontSize: 10.5, color: "#888", marginLeft: "auto" }}>{f.pn}</span>
                    </label>
                  ))}
                </div>
              )}
              {needManual && (
                <div style={{ marginTop: 10, background: "#fff", border: `1px dashed ${C.line}`, padding: 10 }}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: ".08em", color: "#888", marginBottom: 6, fontWeight: 700 }}>Select top assembly</div>
                  <select value={manualTop} onChange={e => setManualTop(e.target.value)} style={{ width: "100%", padding: 6, fontFamily: MONO, fontSize: 11.5 }}>
                    <option value="">— choose —</option>
                    {asmList.map(p => <option key={p.pn} value={p.pn}>{p.pn} — {p.desc}</option>)}
                  </select>
                </div>
              )}
            </div>
          </div>

          {/* Step 3 — order data */}
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={stepNum}>3</div><div style={stepH}>Order Data</div>
            </div>
            <div style={{ borderLeft: `2px solid ${C.line}`, marginLeft: 11, paddingLeft: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <div><label style={lblS}>Work Order</label><input style={inputS} value={wo} onChange={e => setWo(e.target.value)} placeholder="WO-______" /></div>
                <div><label style={lblS}>Serial / Lot</label><input style={inputS} value={sn} onChange={e => setSn(e.target.value)} placeholder="SN-______" /></div>
                <div><label style={lblS}>Program</label><input style={inputS} value={prog} onChange={e => setProg(e.target.value)} /></div>
                <div><label style={lblS}>Issue Date</label><input style={inputS} value={date} onChange={e => setDate(e.target.value)} /></div>
                <div><label style={lblS}>Doc Rev</label><input style={inputS} value={rev} onChange={e => setRev(e.target.value)} /></div>
                <div><label style={lblS}>ECO No.</label><input style={inputS} value={eco} onChange={e => setEco(e.target.value)} /></div>
                <div style={{ gridColumn: "1 / -1" }}><label style={lblS}>Change Description</label><input style={inputS} value={change} onChange={e => setChange(e.target.value)} /></div>
              </div>
              <button style={{ ...btnP, width: "100%", marginTop: 12, padding: 11, fontSize: 13.5, fontWeight: 700, letterSpacing: ".04em" }} onClick={doGenerate}>
                Generate Documents ▸
              </button>

              <div style={{ marginTop: 16, borderTop: `1px solid ${C.line}`, paddingTop: 10 }}>
                <button style={{ ...btn, width: "100%", fontSize: 11 }} onClick={() => setShowSettings(v => !v)}>
                  {showSettings ? "▾" : "▸"} Settings · Adapters · Templates · Project
                </button>
                {showSettings && (
                  <div style={{ marginTop: 8, fontSize: 11, display: "grid", gap: 10 }}>
                    <div style={{ border: `1px solid ${C.line}`, padding: 8, background: "#fff" }}>
                      <b style={{ fontSize: 10.5, letterSpacing: ".05em" }}>OUTPUT PROFILE</b>
                      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                        {Object.values(PROFILES).map(pr => (
                          <button key={pr.id} onClick={() => { setProfile(pr.id); setSheetSize(pr.id === "island" ? "tabloid" : "letter"); }}
                            style={{ ...btn, fontSize: 10.5, flex: 1, background: profile === pr.id ? C.navy : "#fff", color: profile === pr.id ? "#fff" : C.ink, fontWeight: profile === pr.id ? 700 : 400 }}>
                            {pr.label}
                          </button>
                        ))}
                      </div>
                      <div style={{ fontSize: 9.5, color: "#777", marginTop: 6 }}>Profile controls branding, document numbering, column labels, and vocabulary. Generation logic is identical; the Island profile emits Job Travelers + ESP procedures with the ICG title block.</div>
                      {profile === "island" && (
                        <div style={{ marginTop: 8, borderTop: `1px solid ${C.line}`, paddingTop: 8 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
                            <span style={{ fontSize: 10, fontWeight: 700, minWidth: 66 }}>Customer</span>
                            <input value={customerOverride} onChange={e => setCustomerOverride(e.target.value)}
                              placeholder={custDet.value ? custDet.value + " (auto-detected)" : "enter customer"} style={{ ...inputS, flex: 1, fontSize: 10.5, marginTop: 0 }} />
                            {custDet.guessed && !customerOverride && <span style={{ fontSize: 8.5, background: "#FFF200", padding: "1px 4px", fontWeight: 700 }}>auto</span>}
                          </div>
                          <div style={{ fontSize: 10, fontWeight: 700, margin: "6px 0 3px" }}>ESP procedure number per assembly</div>
                          <div style={{ fontSize: 9, color: "#777", marginBottom: 4 }}>Blank = ESP-* (highlighted). Enter the ESP number for each made assembly (stack, stator, etc.).</div>
                          {generated && generated.tops && buildOrder(bom, generated.excluded || {}, generated.tops).map(pn => (
                            <div key={pn} style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                              <span style={{ fontFamily: MONO, fontSize: 9.5, minWidth: 92 }}>{pn}</span>
                              <span style={{ fontSize: 8.5, color: "#888", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(bom.parts[pn] || {}).desc}</span>
                              <input value={espByPn[pn] || ""} onChange={e => setEspByPn(v => ({ ...v, [pn]: e.target.value }))}
                                placeholder="ESP-*" style={{ ...inputS, width: 90, fontSize: 10, marginTop: 0 }} />
                            </div>
                          ))}
                          {!generated && <div style={{ fontSize: 9, color: "#999", fontStyle: "italic" }}>Generate documents to list the assemblies here.</div>}
                        </div>
                      )}
                    </div>
                    <div style={{ border: `1px solid ${C.line}`, padding: 8, background: "#fff" }}>
                      <b style={{ fontSize: 10.5, letterSpacing: ".05em" }}>PROJECT</b>
                      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                        <button style={{ ...btn, fontSize: 10.5 }} onClick={exportProject}>Export project JSON</button>
                        <button style={{ ...btn, fontSize: 10.5 }} onClick={() => projFileRef.current && projFileRef.current.click()}>Import project JSON</button>
                        <input ref={projFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) importProject(e.target.files[0]); e.target.value = ""; }} />
                      </div>
                    </div>
                    <div style={{ border: `1px solid ${C.line}`, padding: 8, background: "#fff" }}>
                      <b style={{ fontSize: 10.5, letterSpacing: ".05em" }}>ROUTING TEMPLATES (JSON)</b>
                      <div style={{ fontSize: 10, color: "#777", margin: "4px 0" }}>Default generation logic is unchanged. Export the built-in templates as editable JSON; import to override by id (or add new). Tokens: {"{PN} {DESC} {ref:regex|fallback} {refPN:regex|fallback}"} · per-op onlyIf regex · autoNumber.</div>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button style={{ ...btn, fontSize: 10.5 }} onClick={() => download("docworks_templates.json", exportTemplatesJSON())}>Export templates</button>
                        <button style={{ ...btn, fontSize: 10.5 }} onClick={() => tplFileRef.current && tplFileRef.current.click()}>Import templates</button>
                        <input ref={tplFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={e => { if (e.target.files[0]) importTemplates(e.target.files[0]); e.target.value = ""; }} />
                        {customDecls.length > 0 && <button style={{ ...btn, fontSize: 10.5, color: "#B03A00", borderColor: "#B03A00" }} onClick={() => { setCustomDecls([]); setCustomTemplates([]); setGenerated(null); }}>Reset to defaults</button>}
                      </div>
                      {customDecls.length > 0 && <div style={{ fontSize: 10, color: "#2E6B3E", marginTop: 4 }}>✓ {customDecls.length} custom template(s) active: {customDecls.map(d => d.id).join(", ")}</div>}
                    </div>
                    <div style={{ border: `1px solid ${C.line}`, padding: 8, background: "#fff" }}>
                      <b style={{ fontSize: 10.5, letterSpacing: ".05em" }}>IMPORT-EDGE ADAPTERS</b>
                      <div style={{ fontSize: 10, color: "#777", margin: "4px 0" }}>OCR / LLM assist BOM import only — document generation is always deterministic.</div>
                      <div style={{ fontSize: 10.5 }}>Tesseract OCR: {(getAdapter("tesseract") || (typeof window !== "undefined" && window.Tesseract)) ? <span style={{ color: "#2E6B3E" }}>● detected</span> : <span style={{ color: "#999" }}>○ not loaded (bundled in desktop build; or register window.docworksAdapters.tesseract)</span>}</div>
                      <div style={{ fontSize: 10.5 }}>PDF text: {getAdapter("pdfText") ? <span style={{ color: "#2E6B3E" }}>● detected</span> : <span style={{ color: "#999" }}>○ not loaded (register window.docworksAdapters.pdfText)</span>}</div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 6 }}>
                        <div><label style={lblS}>Local LLM endpoint</label><input style={inputS} value={llmUrl} onChange={e => setLlmUrl(e.target.value)} /></div>
                        <div><label style={lblS}>Model</label><input style={inputS} value={llmModel} onChange={e => setLlmModel(e.target.value)} /></div>
                      </div>
                      <button style={{ ...btn, fontSize: 10.5, marginTop: 6 }} onClick={async () => {
                        setLlmStatus("testing…");
                        try { const r = await fetch(llmUrl.replace(/\/$/, "") + "/api/tags"); setLlmStatus(r.ok ? "● connected (Ollama-compatible)" : "HTTP " + r.status); }
                        catch (e) { setLlmStatus("○ unreachable — is Ollama running? (" + e.message + ")"); }
                      }}>Test connection</button>
                      {llmStatus && <span style={{ fontSize: 10, marginLeft: 8, color: llmStatus.startsWith("●") ? "#2E6B3E" : "#B03A00" }}>{llmStatus}</span>}
                    </div>
                  </div>
                )}
              </div>
              {unresolvedMissing.length > 0 && (
                <div style={{ fontSize: 10.5, color: "#8A6D00", marginTop: 6, textAlign: "center" }}>
                  ▲ Will generate with {unresolvedMissing.length} unresolved missing-BOM flag{unresolvedMissing.length > 1 ? "s" : ""} shown on the tree
                </div>
              )}
            </div>
          </div>
        </div>

        {/* RIGHT PANE */}
        <div style={{ flex: 1, minWidth: 320, display: "flex", flexDirection: "column", maxHeight: "calc(100vh - 46px)" }}>
          <div style={{ background: C.paper, borderBottom: `1px solid ${C.line}`, display: "flex", alignItems: "center", padding: "0 18px", gap: 2, flexShrink: 0 }}>
            {[["tree", "Family Tree"], ["plist", "Parts List"], ["trav", "Traveler"], ["wi", "Work Instruction"]].map(([id, label]) => (
              <button key={id} disabled={!generated} onClick={() => setTab(id)}
                style={{
                  border: "none", background: "none", padding: "12px 16px", fontSize: 13, fontWeight: 600, cursor: generated ? "pointer" : "default",
                  color: tab === id && generated ? C.navy : "#999", borderBottom: `3px solid ${tab === id && generated ? C.navy : "transparent"}`,
                  opacity: generated ? 1 : .45,
                }}>{label}</button>
            ))}
          </div>
          {generated && (
            <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "6px 16px", borderBottom: `1px solid ${C.line}`, background: "#FCFCFA", flexWrap: "wrap" }}>
              <button onClick={() => setEditMode(v => !v)}
                style={{ border: `1px solid ${editMode ? "#B8860B" : C.navy}`, background: editMode ? "#FFF6DC" : "#fff", color: editMode ? "#8A6D00" : C.navy, padding: "4px 10px", fontSize: 11, fontWeight: 700, borderRadius: 2, cursor: "pointer" }}>
                {editMode ? "✎ Editing ON — click text · right-click rows" : "✎ Edit documents"}
              </button>
              <button onClick={async () => { try { setExportMsg("Opening print view…"); const names = { tree: "Family_Tree", plist: "Parts_List", trav: "Travelers", wi: profile === "island" ? "ESP_Procedures" : "Work_Instructions" }; const ps = tab === "tree" ? (sheetSize === "tabloid" ? "tabloid-landscape" : "letter") : "letter"; await exportPaneAsPDF(docsRef.current, "DocWorks " + names[tab], ps); setExportMsg("Print view opened — choose Save as PDF"); setTimeout(() => setExportMsg(""), 4000); } catch (e) { setExportMsg(e.message); setTimeout(() => setExportMsg(""), 5000); } }}
                style={{ border: `1px solid ${C.navy}`, background: C.navy, color: "#fff", padding: "4px 12px", fontSize: 11, fontWeight: 700, borderRadius: 2, cursor: "pointer" }}>
                ⬇ Save as PDF (exact)
              </button>
              <button onClick={async () => { try { setExportMsg("Exporting…"); const names = { tree: "Family_Tree", plist: "Parts_List", trav: "Travelers", wi: profile === "island" ? "ESP_Procedures" : "Work_Instructions" }; const where = await exportPaneAsWord(docsRef.current, "DocWorks_" + names[tab] + "_" + (generated.tops || []).join("+"), tab === "tree" && sheetSize === "tabloid"); setExportMsg(where === "folder" ? "Saved to folder ✓" : "Downloaded ✓"); setTimeout(() => setExportMsg(""), 3500); } catch (e) { setExportMsg("Export failed: " + e.message); } }}
                style={{ border: `1px solid ${C.navy}`, background: "#fff", color: C.navy, padding: "4px 10px", fontSize: 11, fontWeight: 600, borderRadius: 2, cursor: "pointer" }}>
                ⬇ Word (.docx)
              </button>
              {tab === "plist" && (
                <button onClick={async () => { try { const where = await exportTableAsXlsx(docsRef.current, "DocWorks_Parts_List_" + (generated.tops || []).join("+")); setExportMsg(where === "folder" ? "Saved to folder ✓" : "Downloaded ✓"); setTimeout(() => setExportMsg(""), 3500); } catch (e) { setExportMsg("Export failed: " + e.message); } }}
                  style={{ border: `1px solid ${C.navy}`, background: "#fff", color: C.navy, padding: "4px 10px", fontSize: 11, fontWeight: 600, borderRadius: 2, cursor: "pointer" }}>
                  ⬇ Export Excel (.xlsx)
                </button>
              )}
              <button onClick={async () => { try { const name = await pickExportFolder(); setExportDirName(name); setExportMsg("Export folder set: " + name); setTimeout(() => setExportMsg(""), 3500); } catch (e) { setExportMsg(e.message); setTimeout(() => setExportMsg(""), 5000); } }}
                style={{ border: `1px solid ${C.line}`, background: "#fff", color: "#555", padding: "4px 10px", fontSize: 11, borderRadius: 2, cursor: "pointer" }}>
                📁 {exportDirName ? "Folder: " + exportDirName : "Set export folder…"}
              </button>
              {tab === "tree" && (
                <span style={{ display: "inline-flex", alignItems: "center", gap: 4, marginLeft: 4 }}>
                  <span style={{ fontSize: 10, color: "#777" }}>Sheet:</span>
                  {[["letter", "Letter (multi-sheet)"], ["tabloid", "11×17 (one sheet)"]].map(([id, lbl]) => (
                    <button key={id} onClick={() => setSheetSize(id)}
                      style={{ border: `1px solid ${sheetSize === id ? C.navy : C.line}`, background: sheetSize === id ? C.navy : "#fff", color: sheetSize === id ? "#fff" : "#555", padding: "3px 8px", fontSize: 10.5, fontWeight: sheetSize === id ? 700 : 400, borderRadius: 2, cursor: "pointer" }}>
                      {lbl}
                    </button>
                  ))}
                </span>
              )}
              {editMode && <span style={{ fontSize: 10, color: "#8A6D00" }}>Edits persist until you regenerate — exports capture your edits.</span>}
              {exportMsg && <span style={{ fontSize: 11, fontWeight: 700, color: exportMsg.includes("failed") || exportMsg.includes("needs") ? "#B03A00" : "#2E6B3E" }}>{exportMsg}</span>}
            </div>
          )}
          <div style={{ overflowY: "auto", padding: 22, display: "flex", flexDirection: "column", alignItems: "center", gap: 14, flex: 1 }}>
            {!generated && (
              <div style={{ color: "#999", textAlign: "center", padding: "80px 20px", fontSize: 13 }}>
                <div style={{ fontSize: 38, marginBottom: 10, opacity: .4 }}>⬡</div>
                Drop a BOM (Excel/CSV), paste from a PDF, or use the sample.<br />Then choose what you're building and generate.
              </div>
            )}
            {/* check panel lives OUTSIDE the exportable docsRef so it never appears in exports */}
            {generated && check && (
              <div style={{ marginBottom: 4, width: "100%", maxWidth: 850 }}>
                {check.pass && !check.warns.length && !check.info.length ? (
                  <div style={{ background: "#EAF6EC", border: "1px solid #2E6B3E55", color: "#2E6B3E", padding: "8px 12px", fontSize: 12, fontWeight: 600 }}>✓ Document check passed — routing sequence, hold points, and BOM references verified.</div>
                ) : (
                  <div style={{ border: `1px solid ${C.line}`, background: "#fff", padding: "8px 12px", fontSize: 11.5 }}>
                    <b style={{ fontSize: 11, letterSpacing: ".05em", color: check.errors.length ? "#B03A00" : "#8A6D00" }}>DOCUMENT CHECK — {check.errors.length} error(s), {check.warns.length} warning(s), {check.info.length} note(s) · <span style={{ fontWeight: 400, color: "#999" }}>not included in exports</span></b>
                    {check.errors.map((x, i) => <div key={"e" + i} style={{ color: "#B03A00", marginTop: 3 }}>✕ {x}</div>)}
                    {check.warns.map((x, i) => <div key={"w" + i} style={{ color: "#8A6D00", marginTop: 3 }}>▲ {x}</div>)}
                    {check.info.map((x, i) => <div key={"i" + i} style={{ color: "#666", marginTop: 3 }}>· {x}</div>)}
                  </div>
                )}
              </div>
            )}
            <div ref={docsRef} contentEditable={editMode} suppressContentEditableWarning
              style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 22, width: "100%", outline: editMode ? "2px dashed #B8860B" : "none", outlineOffset: 4, borderRadius: editMode ? 4 : 0 }}>
              {generated && tab === "tree" && <TreeDoc bom={bom} excluded={generated.excluded} tops={generated.tops} cfgName={generated.cfgName} m={m} purchased={generated.purchased || {}} profile={profile} customer={customer} sheetSize={sheetSize} />}
              {generated && tab === "plist" && <PartsListDoc bom={bom} excluded={generated.excluded} tops={generated.tops} cfgName={generated.cfgName} m={m} purchased={generated.purchased || {}} profile={profile} customer={customer} />}
              {generated && tab === "trav" && <TravelerDocs bom={bom} excluded={generated.excluded} tops={generated.tops} m={m} profile={profile} espByPn={espByPn} customer={customer} />}
              {generated && tab === "wi" && <WIDocs bom={bom} excluded={generated.excluded} tops={generated.tops} m={m} profile={profile} espByPn={espByPn} customer={customer} />}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}


ReactDOM.createRoot(document.getElementById("root")).render(<DocWorks />);
