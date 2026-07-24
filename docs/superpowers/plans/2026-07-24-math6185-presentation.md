# MATH6185 Oral Presentation Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and verify an English MATH6185 oral-presentation deck, its PDF backup, an English speaker script, and a bilingual Q&A study guide from the approved report and model evidence.

**Architecture:** Keep numerical evidence in one machine-readable source file and narrative content in a second source file. Generate the PowerPoint with `@oai/artifact-tool`, including speaker notes, and generate the two Word handouts from the same narrative source. Render every final artifact, run structural checks, and compare all displayed values against the final report before delivery.

**Tech Stack:** Node.js ES modules, `@oai/artifact-tool`, Python 3 with `python-docx`, Microsoft PowerPoint COM export, Poppler/PDF tools, bundled presentation and document rendering scripts.

## Global Constraints

- Use an 8–12 minute result-first narrative.
- Create eight core slides plus technical backup slides.
- Audience-facing slide copy and speaker script are English.
- Q&A answers are English with Chinese explanations.
- Do not run AnyLogic live.
- Do not modify the final report or any submitted `.alp` model.
- Use exact values from the final report and the successful 250-replication outputs.
- Use 16:9 slides with navy, white, and teal; use amber or red only for risks and constraints.
- Do not use decorative stock imagery.
- Use at least 50 pt for the deck title, 35 pt for slide titles, 24 pt for subheadings, and 16 pt for body text.
- Fix every unintended overlap, overflow, clipping, wrapping, chart mismatch, and unresolved template text before delivery.

---

## File Structure

**Source and implementation**

- Create: `work/math6185_presentation/source_data.json` — single source of numerical truth and bibliography metadata.
- Create: `work/math6185_presentation/presentation_content.json` — slide copy, speaker notes, optional-shortening flags, and bilingual Q&A content.
- Create: `work/math6185_presentation/validate_sources.mjs` — validates all required numbers, policies, and content entries.
- Create: `work/math6185_presentation/deck.mjs` — builds and exports the PowerPoint with speaker notes.
- Create: `work/math6185_presentation/generate_handouts.py` — builds the speaker-script and Q&A Word files.
- Create: `work/math6185_presentation/export_presentation_pdf.ps1` — exports the final PPTX to PDF through PowerPoint.
- Create: `work/math6185_presentation/validate_outputs.py` — checks final files, page/slide counts, required text, and cross-artifact values.

**Scratch**

- Create: `work/math6185_presentation/tmp/` — artifact-tool workspace, slide renders, document renders, layout exports, montages, and QA logs.

**Final deliverables**

- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pdf`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script.docx`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide.docx`

---

### Task 1: Initialise Version Control and Lock the Evidence

**Files:**
- Create: `.git/`
- Create: `work/math6185_presentation/source_data.json`
- Create: `work/math6185_presentation/presentation_content.json`
- Create: `work/math6185_presentation/validate_sources.mjs`
- Existing evidence: `outputs/zw1f25_MATH6185_submission/zw1f25_MATH6185_Report.pdf`
- Existing evidence: `outputs/zw1f25_MATH6185_submission/zw1f25_MATH6185_Q2_FAS.alp`
- Existing evidence: `outputs/zw1f25_MATH6185_submission/zw1f25_MATH6185_Q3_WorkloadPolicy_alpha035.alp`
- Existing evidence: `outputs/zw1f25_MATH6185_submission/zw1f25_MATH6185_Q4_OCR_eta020_L0_H3.alp`

**Interfaces:**
- Produces: `source_data.json` with keys `system`, `experiment`, `q2`, `q3`, `q4`, `comparison`, and `references`.
- Produces: `presentation_content.json` with arrays `slides` and `qa`.
- Produces: a validator that exits 0 only when every required value and content section is present.

- [ ] **Step 1: Initialise Git and commit the approved design**

Run:

```powershell
git init -b main
git add docs/superpowers/specs/2026-07-24-math6185-presentation-design.md
git commit -m "docs: approve MATH6185 presentation design"
```

Expected: a new `main` repository with one commit containing only the approved design.

- [ ] **Step 2: Create the numerical source of truth**

Create `source_data.json` with these exact values:

```json
{
  "system": {
    "urgentArrivalRatePerDay": 3.69,
    "routineArrivalRatePerDay": 2.19,
    "slotsPerDay": 6,
    "nominalLoad": 0.98,
    "overallMeanConstraintDays": 5
  },
  "experiment": {
    "warmupServed": 1000,
    "stopAtArrivals": 50000,
    "replications": 250,
    "confidenceLevel": 0.95,
    "relativeHalfWidthTarget": 0.05
  },
  "q2": {
    "urgentMean": 4.147618857314027,
    "routineMean": 4.14819901597308,
    "overallMean": 4.147829551134296,
    "overallCI": [3.9776382383429376, 4.318020863925655],
    "relativeHalfWidth": 0.04103141430795218,
    "warmup": {
      "500": {"mean": 4.1332073379295675, "ci": [3.964507168938908, 4.3019075069202275]},
      "1000": {"mean": 4.147829551134296, "ci": [3.9776382383429376, 4.318020863925655]},
      "2000": {"mean": 4.171597031126426, "ci": [3.998668591384816, 4.344525470868036]}
    }
  },
  "q3": {
    "alpha1": 0,
    "alpha2": 0.35,
    "urgentMean": 2.123416974280667,
    "urgentCI": [2.074938008716257, 2.171895939845077],
    "routineMean": 9.052809827419521,
    "routineCI": [8.62113318254977, 9.484486472289273],
    "overallMean": 4.702753976854765,
    "overallCI": [4.512624489159033, 4.892883464550496],
    "relativeHalfWidth": 0.04042939278377726,
    "urgentReductionDays": 2.02441257685336,
    "urgentReductionPercent": 48.805339
  },
  "q4": {
    "eta": 0.2,
    "L": 0,
    "H": 3,
    "urgentMean": 4.0090591696293085,
    "urgentCI": [3.862104540343827, 4.15601379891479],
    "routineMean": 4.573754083738393,
    "routineCI": [4.425944769671173, 4.721563397805613],
    "overallMean": 4.219449679893367,
    "overallCI": [4.072177031422903, 4.3667223283638315],
    "relativeHalfWidth": 0.03490328351876109,
    "utilization": 0.9791112781532182,
    "utilizationCI": [0.9785772006178842, 0.9796453556885522],
    "meanDaysSaved": 0.5146292846991197,
    "urgentReductionDays": 0.1387703815049877,
    "urgentReductionPercent": 3.345609
  }
}
```

Add the nine report references under `references`, preserving author, year, title, journal, volume, issue, pages, and DOI. Use the full Patrick and Puterman (2007) title including its subtitle.

- [ ] **Step 3: Create the narrative source**

Create `presentation_content.json` with:

- Fourteen slide records: eight core and six backup.
- Each record contains `id`, `title`, `purpose`, `visibleCopy`, `speakerNotes`, `shortVersionOmissions`, and `sources`.
- The six backup slides are: full confidence intervals, warm-up sensitivity, Q3 parameter screen, Q4 policy screen, AnyLogic implementation map, and literature/AI disclosure.
- At least 18 Q&A records covering problem motivation, FAS validation, warm-up, stopping rule, replication design, confidence intervals, Q3 rule, Q3 parameter choice, routine fairness, Q4 mechanism, Q4 parameter choice, utilisation, H=2 versus H=3, event-equivalent screening, literature, limitations, original thinking, and AI assistance.
- Each Q&A record contains `question`, `answerEnglish`, `explanationChinese`, and `evidence`.

- [ ] **Step 4: Write the evidence validator**

Implement `validate_sources.mjs` with:

```javascript
import fs from "node:fs";
import assert from "node:assert/strict";

const data = JSON.parse(fs.readFileSync(new URL("./source_data.json", import.meta.url), "utf8"));
const content = JSON.parse(fs.readFileSync(new URL("./presentation_content.json", import.meta.url), "utf8"));

assert.equal(data.experiment.replications, 250);
assert.equal(data.experiment.stopAtArrivals, 50000);
assert.equal(data.q3.alpha2, 0.35);
assert.equal(data.q4.eta, 0.2);
assert.equal(data.q4.L, 0);
assert.equal(data.q4.H, 3);
assert.ok(data.q3.overallCI[1] < 5);
assert.equal(content.slides.filter((s) => s.kind === "core").length, 8);
assert.equal(content.slides.filter((s) => s.kind === "backup").length, 6);
assert.ok(content.qa.length >= 18);
assert.ok(content.slides.every((s) => s.title && s.speakerNotes && Array.isArray(s.sources)));
assert.ok(content.qa.every((q) => q.question && q.answerEnglish && q.explanationChinese && q.evidence));

console.log("SOURCE VALIDATION PASS");
```

- [ ] **Step 5: Run the source validator**

Run:

```powershell
node work/math6185_presentation/validate_sources.mjs
```

Expected:

```text
SOURCE VALIDATION PASS
```

- [ ] **Step 6: Commit the evidence layer**

Run:

```powershell
git add work/math6185_presentation/source_data.json work/math6185_presentation/presentation_content.json work/math6185_presentation/validate_sources.mjs
git commit -m "data: lock MATH6185 presentation evidence"
```

Expected: a second commit containing only the source data, narrative content, and validator.

---

### Task 2: Build the PowerPoint Deck

**Files:**
- Create: `work/math6185_presentation/deck.mjs`
- Create: `work/math6185_presentation/tmp/`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx`
- Read: `work/math6185_presentation/source_data.json`
- Read: `work/math6185_presentation/presentation_content.json`

**Interfaces:**
- Consumes: validated numerical and narrative JSON.
- Produces: `buildDeck(data, content): Promise<Presentation>`.
- Produces: a 14-slide PPTX with speaker notes on every slide.

- [ ] **Step 1: Read the presentation authoring requirements**

Read completely:

```text
D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations\references\content-rules.md
D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations\artifact_tool\API_QUICK_START.md
D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations\artifact_tool\api\API_DOCS.md
D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations\artifact_tool\api\references\speaker-notes.spec.md
```

Expected: the deck authoring code uses `@oai/artifact-tool` only and sets notes through `slide.speakerNotes.textFrame.setText(...)`.

- [ ] **Step 2: Initialise the artifact-tool workspace**

Run:

```powershell
$skill='D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
node "$skill\container_tools\setup_artifact_tool_workspace.mjs" --workspace "work\math6185_presentation\tmp"
```

Expected: `work/math6185_presentation/tmp/package.json` and resolvable `@oai/artifact-tool`.

- [ ] **Step 3: Implement the deck helpers**

In `deck.mjs`, define these interfaces:

```javascript
function addTitle(slide, title, subtitle = "") {}
function addFooter(slide, sourceText, slideNumber) {}
function addSpeakerNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
}
function addMeanComparisonChart(slide, data) {}
function addModelFlow(slide) {}
function addQ4RebookingFlow(slide) {}
function addConfidenceTable(slide, rows) {}
async function buildDeck(data, content) {}
```

Use these colours:

```javascript
const COLORS = {
  navy: "#17324D",
  teal: "#007F82",
  paleTeal: "#DCEEEF",
  white: "#FFFFFF",
  ink: "#18222C",
  grey: "#687783",
  lightGrey: "#EEF2F4",
  amber: "#D68B00",
  red: "#B23A48"
};
```

Use a 16:9 page size, a single sans-serif family available on Windows, and consistent title/footer positions across all slides.

- [ ] **Step 4: Build the eight core slides**

Implement the eight core slides exactly as defined in the approved design. Required visual evidence:

- Slide 3 comparison chart uses:
  - FAS: urgent 4.148, routine 4.148, overall 4.148
  - Q3: urgent 2.123, routine 9.053, overall 4.703
  - Q4: urgent 4.009, routine 4.574, overall 4.219
- Slide 4 displays 3.69/day, 2.19/day, six slots/day, 1,000 served warm-up, 50,000-arrival stop, 250 replications, and 95% CIs.
- Slide 5 displays Q2 mean 4.148 and CI [3.978, 4.318].
- Slide 6 displays `alpha1 = 0`, `alpha2 = 0.35`, urgent 2.123, routine 9.053, overall 4.703, and 48.8%.
- Slide 7 displays `eta = 0.2`, `L = 0`, `H = 3`, urgent 4.009, routine 4.574, overall 4.219, utilisation 97.911%, and 0.515 days saved.
- Slide 8 recommends Q3 while explicitly naming the routine-wait service-level risk.

- [ ] **Step 5: Build the six backup slides**

Implement all six backup slides and mark them visually as `BACKUP` without including timing instructions in audience-facing copy. Use compact tables but keep body text at 16 pt or larger.

- [ ] **Step 6: Attach English speaker notes**

For every slide:

```javascript
const noteText = content.slides.find((item) => item.id === slideId).speakerNotes;
addSpeakerNotes(slide, noteText);
```

Expected: all 14 slides expose notes in PowerPoint Presenter View.

- [ ] **Step 7: Export the PPTX**

Export to:

```text
outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx
```

Run:

```powershell
node work/math6185_presentation/deck.mjs
```

Expected: a non-empty PPTX containing 14 slides.

- [ ] **Step 8: Run structural slide checks**

Run:

```powershell
$skill='D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skill\container_tools\slides_test.py" "outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Presentation.pptx"
```

Expected: no overflow or out-of-bounds findings.

- [ ] **Step 9: Commit the deck implementation**

Run:

```powershell
git add work/math6185_presentation/deck.mjs outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx
git commit -m "feat: build MATH6185 oral presentation deck"
```

---

### Task 3: Build the Speaker Script and Q&A Guide

**Files:**
- Create: `work/math6185_presentation/generate_handouts.py`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script.docx`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide.docx`
- Read: `work/math6185_presentation/presentation_content.json`

**Interfaces:**
- Consumes: `slides[*].speakerNotes`, `slides[*].shortVersionOmissions`, and `qa[*]`.
- Produces: two styled Word documents with a table of contents-like opening guide and consistent headings.

- [ ] **Step 1: Read the document skill requirements**

Read completely:

```text
D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\documents\26.715.12143\skills\documents\SKILL.md
```

- [ ] **Step 2: Implement the document generator**

Create `generate_handouts.py` with these functions:

```python
def configure_document(document, title: str) -> None: ...
def add_script_slide(document, slide: dict) -> None: ...
def add_qa_entry(document, entry: dict, number: int) -> None: ...
def build_speaker_script(content: dict, output_path: str) -> None: ...
def build_qa_guide(content: dict, output_path: str) -> None: ...
```

Formatting requirements:

- A4 pages with 2 cm margins.
- Aptos or Calibri, 11 pt body text.
- Clear Heading 1 and Heading 2 hierarchy.
- Speaker script starts with a one-page delivery guide and then one section per core slide.
- Optional shorter-version sentences appear in muted grey and begin with `[Optional if time is short]`.
- Q&A guide uses one numbered entry per question, with the English answer first and a shaded Chinese explanation below.
- Each Q&A entry ends with an `Evidence:` line.

- [ ] **Step 3: Generate both Word files**

Run:

```powershell
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' work/math6185_presentation/generate_handouts.py
```

Expected: both DOCX files exist and are larger than 20 KB.

- [ ] **Step 4: Commit the handout implementation**

Run:

```powershell
git add work/math6185_presentation/generate_handouts.py outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script.docx outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide.docx
git commit -m "feat: add MATH6185 speaker script and Q&A guide"
```

---

### Task 4: Export and Visually Verify All Artifacts

**Files:**
- Create: `work/math6185_presentation/export_presentation_pdf.ps1`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pdf`
- Create: `work/math6185_presentation/tmp/rendered_slides/`
- Create: `work/math6185_presentation/tmp/rendered_script/`
- Create: `work/math6185_presentation/tmp/rendered_qa/`
- Create: `work/math6185_presentation/tmp/qa/visual_review.txt`

**Interfaces:**
- Consumes: the final PPTX and DOCX files.
- Produces: the final PDF and a completed visual QA ledger.

- [ ] **Step 1: Implement PowerPoint PDF export**

Create `export_presentation_pdf.ps1`:

```powershell
$pptx = (Resolve-Path 'outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Presentation.pptx').Path
$pdf = Join-Path (Split-Path $pptx) 'zw1f25_MATH6185_Presentation.pdf'
$app = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $app.Presentations.Open($pptx, $true, $false, $false)
    try {
        $presentation.SaveAs($pdf, 32)
    } finally {
        $presentation.Close()
    }
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
```

- [ ] **Step 2: Export the PDF**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File work/math6185_presentation/export_presentation_pdf.ps1
```

Expected: a 14-page PDF matching the 14-slide deck.

- [ ] **Step 3: Render every slide**

Run:

```powershell
$skill='D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skill\container_tools\render_slides.py" "outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Presentation.pptx"
```

Move the resulting slide PNGs into `work/math6185_presentation/tmp/rendered_slides/`.

- [ ] **Step 4: Create a deck montage**

Run:

```powershell
$skill='D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skill\container_tools\create_montage.py" --input_dir "work\math6185_presentation\tmp\rendered_slides" --output_file "work\math6185_presentation\tmp\qa\deck_montage.png"
```

- [ ] **Step 5: Inspect every slide at full size**

For slides 1–14, inspect the PNG and record one of:

```text
PASS — no clipping, overlap, wrapping, data mismatch, or unreadable text.
FIXED — describe the specific defect and the corrected slide element.
```

Store the entries in `work/math6185_presentation/tmp/qa/visual_review.txt`. Rebuild and re-render after every fix.

- [ ] **Step 6: Render the two Word documents**

Use the document skill's `render_docx.py` to render each DOCX into its own directory:

```powershell
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' 'D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\documents\26.715.12143\skills\documents\render_docx.py' --input "outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Speaker_Script.docx" --output_dir "work\math6185_presentation\tmp\rendered_script"
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' 'D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\documents\26.715.12143\skills\documents\render_docx.py' --input "outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_QA_Guide.docx" --output_dir "work\math6185_presentation\tmp\rendered_qa"
```

Expected: every page renders to PNG without a LibreOffice or font error.

- [ ] **Step 7: Inspect every handout page**

Inspect all rendered pages and fix:

- clipped text or tables;
- orphaned headings;
- unreadable Chinese text;
- inconsistent spacing;
- broken page numbers;
- accidental blank pages.

Append the page-level result to `visual_review.txt`.

- [ ] **Step 8: Commit export and visual QA**

Run:

```powershell
git add work/math6185_presentation/export_presentation_pdf.ps1 work/math6185_presentation/tmp/qa/visual_review.txt outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pdf
git commit -m "test: render and visually verify presentation package"
```

---

### Task 5: Run Final Cross-Artifact Validation and Package the Handoff

**Files:**
- Create: `work/math6185_presentation/validate_outputs.py`
- Verify: all four final deliverables

**Interfaces:**
- Consumes: final PPTX, PDF, speaker script, Q&A guide, source data, and narrative content.
- Produces: exit code 0 and `FINAL OUTPUT VALIDATION PASS`.

- [ ] **Step 1: Implement the final validator**

Create `validate_outputs.py` with these checks:

```python
from pathlib import Path
from zipfile import ZipFile
from pypdf import PdfReader
from docx import Document

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "outputs" / "zw1f25_MATH6185_presentation"
PPTX = OUT / "zw1f25_MATH6185_Presentation.pptx"
PDF = OUT / "zw1f25_MATH6185_Presentation.pdf"
SCRIPT = OUT / "zw1f25_MATH6185_Speaker_Script.docx"
QA = OUT / "zw1f25_MATH6185_QA_Guide.docx"

for path in (PPTX, PDF, SCRIPT, QA):
    assert path.exists() and path.stat().st_size > 20_000, path

with ZipFile(PPTX) as zf:
    slide_files = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    assert len(slide_files) == 14
    ppt_text = " ".join(zf.read(n).decode("utf-8", errors="ignore") for n in slide_files)

pdf = PdfReader(str(PDF))
assert len(pdf.pages) == 14

script_text = "\n".join(p.text for p in Document(SCRIPT).paragraphs)
qa_text = "\n".join(p.text for p in Document(QA).paragraphs)

for token in ("4.148", "2.123", "9.053", "4.703", "4.009", "4.574", "4.219", "97.911", "0.515"):
    assert token in ppt_text or token in script_text, token

assert "250 independent replications" in script_text
assert "50,000" in script_text
assert "人工智能" in qa_text or "AI" in qa_text
assert len([p for p in Document(QA).paragraphs if p.text.startswith("Q")]) >= 18

print("FINAL OUTPUT VALIDATION PASS")
```

- [ ] **Step 2: Run every automated check**

Run:

```powershell
node work/math6185_presentation/validate_sources.mjs
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' work/math6185_presentation/validate_outputs.py
$skill='D:\Users\25470\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
& 'D:\Users\25470\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skill\container_tools\slides_test.py" "outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Presentation.pptx"
```

Expected:

```text
SOURCE VALIDATION PASS
FINAL OUTPUT VALIDATION PASS
```

and no slide overflow findings.

- [ ] **Step 3: Confirm source files were not modified**

Run:

```powershell
Get-FileHash outputs\zw1f25_MATH6185_submission\zw1f25_MATH6185_Report.pdf -Algorithm SHA256
Get-FileHash outputs\zw1f25_MATH6185_submission\zw1f25_MATH6185_Q2_FAS.alp -Algorithm SHA256
Get-FileHash outputs\zw1f25_MATH6185_submission\zw1f25_MATH6185_Q3_WorkloadPolicy_alpha035.alp -Algorithm SHA256
Get-FileHash outputs\zw1f25_MATH6185_submission\zw1f25_MATH6185_Q4_OCR_eta020_L0_H3.alp -Algorithm SHA256
```

Expected hashes:

```text
Report  7DD574347034EF6B055C1AD60908C7D1BA828D6898A7E53991A1B8B0BAEA942C
Q2      AA8F471BECA3C1AAF987561FBA1EA259CCBFC3484BB08E733C126995270E2C76
Q3      21AB826240270C02A8B703F55B92ECA402798F5E535FAA73FC9C6EA90E086BFE
Q4      78C83C924107C12B98E9C85834138FFA4E63AEBF0E79D26F1D8EA31C8FDB830F
```

- [ ] **Step 4: Commit the final validator and verified deliverables**

Run:

```powershell
git add work/math6185_presentation/validate_outputs.py outputs/zw1f25_MATH6185_presentation
git commit -m "chore: finalise verified MATH6185 presentation package"
git status --short
```

Expected: clean status for all files placed under version control by this plan.

- [ ] **Step 5: Deliver the package**

Provide clickable links to:

```text
outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx
outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pdf
outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script.docx
outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide.docx
```

State the core slide count, backup slide count, supported talk length, and the completed QA checks. Do not claim that AnyLogic was rerun during presentation preparation.
