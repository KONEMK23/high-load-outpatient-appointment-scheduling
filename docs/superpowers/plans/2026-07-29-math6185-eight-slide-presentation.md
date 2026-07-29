# MATH6185 Eight-Slide Presentation Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create an eight-slide copy of the existing MATH6185 presentation, a matching eight-page PDF, and an 18-question bilingual Q&A guide focused only on Slides 1-8.

**Architecture:** Use the existing 14-slide PPTX as the sole visual source and duplicate source Slides 1-8 through the presentation template-following workflow. Create the PDF by retaining source PDF pages 1-8, and generate the Q&A guide from one structured bilingual data set with deterministic validation. Compare slide/package structure and rendered pages against the source before delivery.

**Tech Stack:** Bundled Node.js, `@oai/artifact-tool`, presentation template-following helpers, bundled Python, `python-docx`, `pypdf`, Poppler, OOXML ZIP inspection, Pillow image comparison.

## Global Constraints

- Preserve the original 14-slide PPTX, 14-page PDF, and full bilingual Q&A guide unchanged.
- The new presentation must contain exact copies of source Slides 1-8 in the same order.
- Preserve visible content, theme, masters, layouts, fonts, charts, notes, citations, footers, and slide numbering for Slides 1-8.
- The new PDF must contain exactly eight pages matching the retained presentation slides.
- The focused Q&A guide must contain one quick-location table, eight slide sections, and exactly 18 bilingual questions.
- Every question must include an English question, Chinese question, concise English answer, faithful Chinese answer, listening keywords, and one number or mechanism cue.
- Answers must be suitable for a 20-40 second oral response.
- Omit backup-only topics: detailed H = 2 versus H = 3 comparison, event-equivalent screening, full confidence-interval table, and a separate AI-disclosure question.
- Use only the validated numerical results already present in the deck and successful 250-replication outputs.

---

## File Structure

**Implementation**

- Create: `work/math6185_eight_slide_package/build_eight_slide_deck.mjs` - imports and re-exports the template starter through `@oai/artifact-tool`.
- Create: `work/math6185_eight_slide_package/build_focused_qa.py` - structured bilingual Q&A data and DOCX generator.
- Create: `work/math6185_eight_slide_package/build_eight_page_pdf.py` - retains source PDF pages 1-8.
- Create: `work/math6185_eight_slide_package/validate_package.py` - validates all three deliverables and source preservation.
- Create: `work/math6185_eight_slide_package/test_focused_qa.py` - checks Q&A coverage, ordering, and numerical anchors.

**Temporary QA**

- Create: `work/math6185_eight_slide_package/tmp/` - artifact-tool workspace, template inventory, frame map, starter deck, renders, image-diff outputs, and QA logs.

**Final deliverables**

- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pptx`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pdf`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide_PPT_Focused_Bilingual.docx`

---

### Task 1: Build and validate the focused bilingual Q&A guide

**Files:**
- Create: `work/math6185_eight_slide_package/test_focused_qa.py`
- Create: `work/math6185_eight_slide_package/build_focused_qa.py`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide_PPT_Focused_Bilingual.docx`

**Interfaces:**
- Consumes: a `QUESTIONS: list[dict]` constant containing exactly 18 records.
- Produces: `validate_question_data() -> None` and `build_document(output_path: Path) -> None`.
- Each record contains `slide`, `question_en`, `question_zh`, `answer_en`, `answer_zh`, `keywords`, and `cue`.

- [ ] **Step 1: Write failing coverage tests**

```python
from build_focused_qa import QUESTIONS, validate_question_data

validate_question_data()
assert len(QUESTIONS) == 18
assert [item["slide"] for item in QUESTIONS] == [
    1, 1, 2, 2, 3, 3, 4, 4, 4, 5, 5, 6, 6, 6, 7, 7, 8, 8
]
assert all(set(item) == {
    "slide", "question_en", "question_zh", "answer_en",
    "answer_zh", "keywords", "cue"
} for item in QUESTIONS)
for token in ("0.98", "250", "50,000", "1,000", "4.148", "2.123",
              "9.053", "4.703", "4.009", "97.911%", "0.515"):
    assert any(token in item["answer_en"] or token in item["cue"] for item in QUESTIONS)
```

- [ ] **Step 2: Run the tests and confirm the builder is absent**

Run:

```powershell
& $PYTHON work\math6185_eight_slide_package\test_focused_qa.py
```

Expected: import failure for `build_focused_qa`.

- [ ] **Step 3: Implement the structured 18-question data set**

Define exactly these slide/question assignments:

```python
QUESTION_TITLES = [
    (1, "What is the central decision problem?"),
    (1, "What is the main result, and why is Q3 recommended?"),
    (2, "How is the nominal load of 0.98 calculated, and what does it imply?"),
    (2, "How did the literature shape the central trade-off?"),
    (3, "How do Q2, Q3 and Q4 compare?"),
    (3, "Why recommend Q3 despite its effect on routine patients?"),
    (4, "Why use the same experiment design and stop at 50,000 arrivals?"),
    (4, "Why use a warm-up of 1,000 served patients?"),
    (4, "Why use 250 independent replications and 95% confidence intervals?"),
    (5, "Why is FAS a useful benchmark?"),
    (5, "How was the Q2 benchmark validated?"),
    (6, "How does Q3 work?"),
    (6, "Why was alpha2 = 0.35 selected?"),
    (6, "Why does routine waiting rise to 9.053 days?"),
    (7, "How does Q4 work, and what do eta, L and H mean?"),
    (7, "Why is Q4's urgent improvement small despite 97.911% utilisation?"),
    (8, "Would Q3 be implemented unchanged?"),
    (8, "What are the main limitations and the next useful experiment?"),
]
```

Use the approved values: Q2 overall `4.148` with CI `[3.978, 4.318]`; Q3 urgent `2.123`, routine `9.053`, overall `4.703`, upper CI `4.893`, `alpha1 = 0`, `alpha2 = 0.35`; Q4 urgent `4.009`, routine `4.574`, overall `4.219`, utilisation `97.911%`, days saved `0.515`, `eta = 0.2`, `L = 0`, `H = 3`.

- [ ] **Step 4: Implement the compact-reference-guide DOCX**

Use the `compact_reference_guide` preset with:

```python
PAGE = {"width": 8.5, "height": 11.0, "margin": 1.0}
BODY = {"font": "Calibri", "size": 11, "after": 6, "line": 1.25}
H1 = {"size": 16, "color": "2E74B5", "before": 18, "after": 10}
H2 = {"size": 13, "color": "2E74B5", "before": 14, "after": 7}
H3 = {"size": 12, "color": "1F4D78", "before": 10, "after": 5}
TABLE_WIDTH_DXA = 9360
TABLE_INDENT_DXA = 120
TABLE_CELL_MARGINS_DXA = {"top": 80, "bottom": 80, "start": 120, "end": 120}
```

Create a one-page opening block titled `MATH6185 PPT-Focused Bilingual Q&A Guide`, followed by a quick-location table with columns `Slide`, `Listen for`, and `Questions`. Add eight Heading 1 slide sections. For every question, render the seven fields in the approved order and keep the question heading with the answer that follows.

- [ ] **Step 5: Run coverage tests and generate the document**

Run:

```powershell
& $PYTHON work\math6185_eight_slide_package\test_focused_qa.py
& $PYTHON work\math6185_eight_slide_package\build_focused_qa.py
```

Expected: coverage tests pass and the target DOCX exists with non-zero size.

- [ ] **Step 6: Run structural DOCX checks**

Open the generated DOCX with `python-docx` and require:

```python
assert len(question_headings) == 18
assert len(slide_headings) == 8
assert len(document.tables) == 1
assert [row.cells[0].text for row in document.tables[0].rows[1:]] == [
    "Slide 1", "Slide 2", "Slide 3", "Slide 4",
    "Slide 5", "Slide 6", "Slide 7", "Slide 8"
]
```

Expected: all checks pass.

### Task 2: Create the exact eight-slide presentation copy

**Files:**
- Create: `work/math6185_eight_slide_package/tmp/template-frame-map.json`
- Create: `work/math6185_eight_slide_package/tmp/template-audit.txt`
- Create: `work/math6185_eight_slide_package/tmp/deviation-log.txt`
- Create: `work/math6185_eight_slide_package/build_eight_slide_deck.mjs`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pptx`

**Interfaces:**
- Consumes: the source 14-slide PPTX and a frame map for source Slides 1-8.
- Produces: an artifact-tool-exported PPTX with exactly eight slides.

- [ ] **Step 1: Initialise the artifact-tool workspace**

Run:

```powershell
& $NODE "$SKILL_DIR\container_tools\setup_artifact_tool_workspace.mjs" `
  --workspace "work\math6185_eight_slide_package\tmp"
```

Expected: the temporary workspace resolves `@oai/artifact-tool`.

- [ ] **Step 2: Inspect all 14 source slides**

Run:

```powershell
& $NODE "$SKILL_DIR\template_following_scripts\inspect_template_deck.mjs" `
  --workspace "work\math6185_eight_slide_package\tmp" `
  --pptx "$SOURCE_PPTX"
```

Review all slide PNGs, layout JSON, `template-inspect.ndjson`, media, font evidence, and `template-manifest.json`. Record the source hierarchy, slide dimensions, notes presence, fonts, footer behavior, and any inherited placeholders in `template-audit.txt`.

- [ ] **Step 3: Create the preserve-only frame map**

Write:

```json
{
  "outputSlides": [
    {"outputSlide": 1, "sourceSlide": 1, "narrativeRole": "decision and result", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 2, "sourceSlide": 2, "narrativeRole": "capacity pressure and literature", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 3, "sourceSlide": 3, "narrativeRole": "cross-policy comparison", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 4, "sourceSlide": 4, "narrativeRole": "experiment design", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 5, "sourceSlide": 5, "narrativeRole": "Q2 benchmark", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 6, "sourceSlide": 6, "narrativeRole": "Q3 policy", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 7, "sourceSlide": 7, "narrativeRole": "Q4 policy", "reuseMode": "duplicate-slide", "editTargets": []},
    {"outputSlide": 8, "sourceSlide": 8, "narrativeRole": "recommendation", "reuseMode": "duplicate-slide", "editTargets": []}
  ],
  "omittedSourceSlides": [
    {"sourceSlide": 9, "reason": "backup slide excluded by approved scope"},
    {"sourceSlide": 10, "reason": "backup slide excluded by approved scope"},
    {"sourceSlide": 11, "reason": "backup slide excluded by approved scope"},
    {"sourceSlide": 12, "reason": "backup slide excluded by approved scope"},
    {"sourceSlide": 13, "reason": "backup slide excluded by approved scope"},
    {"sourceSlide": 14, "reason": "backup slide excluded by approved scope"}
  ]
}
```

Write `deviation-log.txt` with `No intentional visual or content deviations from source Slides 1-8.`

- [ ] **Step 4: Build the eight-slide starter deck**

Run:

```powershell
& $NODE "$SKILL_DIR\template_following_scripts\prepare_template_starter_deck.mjs" `
  --workspace "work\math6185_eight_slide_package\tmp" `
  --pptx "$SOURCE_PPTX" `
  --map "work\math6185_eight_slide_package\tmp\template-frame-map.json" `
  --out "work\math6185_eight_slide_package\tmp\template-starter.pptx" `
  --preview-dir "work\math6185_eight_slide_package\tmp\template-starter-preview" `
  --layout-dir "work\math6185_eight_slide_package\tmp\template-starter-layout" `
  --contact-sheet "work\math6185_eight_slide_package\tmp\template-starter-contact-sheet.png"
```

Expected: an eight-slide starter whose map validates with no edit targets.

- [ ] **Step 5: Import and export the starter through artifact-tool**

Implement:

```javascript
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const presentation = await PresentationFile.importPptx(
  await FileBlob.load(process.env.STARTER_PPTX),
);
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(process.env.FINAL_PPTX);
```

Export every slide to PNG and layout JSON under the temporary QA folder before exporting the final PPTX.

- [ ] **Step 6: Run template fidelity and slide overflow checks**

Run `check_template_fidelity.mjs` against the starter and final PPTX, then run `slides_test.py` against the final PPTX. Expected: no unplanned edits, no empty inherited placeholders, no overflow, and no out-of-bounds objects.

### Task 3: Create and verify the matching eight-page PDF

**Files:**
- Create: `work/math6185_eight_slide_package/build_eight_page_pdf.py`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pdf`

**Interfaces:**
- Consumes: source 14-page PDF.
- Produces: an eight-page PDF containing source pages 1-8 without rasterisation.

- [ ] **Step 1: Write the page-subset builder**

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader(str(source))
assert len(reader.pages) == 14
writer = PdfWriter()
for page in reader.pages[:8]:
    writer.add_page(page)
with target.open("wb") as stream:
    writer.write(stream)
```

- [ ] **Step 2: Generate and reopen the PDF**

Run the builder, reopen the output with `PdfReader`, and require exactly eight pages with the same media boxes as source pages 1-8.

- [ ] **Step 3: Render all PDF pages**

Use bundled Poppler:

```powershell
pdftoppm -png -r 144 "$OUTPUT_PDF" "work\math6185_eight_slide_package\tmp\pdf\page"
```

Expected: eight readable PNGs with no clipping, missing glyphs, or black boxes.

### Task 4: Run final cross-artifact and visual validation

**Files:**
- Create: `work/math6185_eight_slide_package/validate_package.py`
- Verify: all three final deliverables and all three original source files.

**Interfaces:**
- Consumes: source and output PPTX/PDF/DOCX plus rendered page images.
- Produces: exit code `0` and `FINAL EIGHT-SLIDE PACKAGE VALIDATION PASS`.

- [ ] **Step 1: Record source hashes before generation**

Use SHA-256 hashes for the source PPTX, PDF, and full bilingual Q&A guide. Store them in the temporary QA ledger.

- [ ] **Step 2: Implement PPTX package checks**

For source and output PPTX packages:

```python
assert len(output_slide_parts) == 8
assert len(source_slide_parts) == 14
assert output_slide_texts == source_slide_texts[:8]
assert output_note_texts == source_note_texts[:8]
assert output_slide_relationship_targets == source_slide_relationship_targets[:8]
```

Also verify slide size, slide order, notes presence, chart/media counts, master/layout references, theme names, footer text, and slide-number text.

- [ ] **Step 3: Implement PDF and DOCX checks**

Require:

```python
assert len(PdfReader(str(output_pdf)).pages) == 8
assert len(question_headings) == 18
assert len(slide_headings) == 8
assert len(qa_document.tables) == 1
```

Check every required value against the approved numerical anchors and reject backup-only phrases `event-equivalent`, `H = 2`, `full confidence interval table`, and `AI assistance`.

- [ ] **Step 4: Compare renders**

Render source PPTX Slides 1-8, final PPTX Slides 1-8, source PDF pages 1-8, and final PDF pages 1-8. Use Pillow `ImageChops.difference` after equal-size conversion and require exact equality for the source/final PDF page pairs. For PPTX source/final pairs, inspect any non-zero difference and require that no visible content, geometry, colour, font, or crop changed.

- [ ] **Step 5: Render and inspect the focused DOCX**

Run:

```powershell
& $PYTHON "$DOC_SKILL_DIR\render_docx.py" "$OUTPUT_DOCX" `
  --output_dir "work\math6185_eight_slide_package\tmp\docx-render"
```

Inspect every page for Chinese glyphs, clipped text, broken tables, orphaned question headings, inconsistent spacing, and accidental blank pages. If LibreOffice is unavailable, record that specific fallback and complete structural OOXML validation without claiming visual DOCX QA.

- [ ] **Step 6: Run final verification and confirm source preservation**

Run:

```powershell
& $PYTHON work\math6185_eight_slide_package\test_focused_qa.py
& $PYTHON work\math6185_eight_slide_package\validate_package.py
git diff --check
```

Recompute the three source SHA-256 hashes and require equality with the values recorded in Step 1. Expected:

```text
Focused Q&A coverage: PASS
PPTX slide count and source preservation: PASS
PDF page count and render identity: PASS
DOCX structure and numerical anchors: PASS
Original source hashes unchanged: PASS
FINAL EIGHT-SLIDE PACKAGE VALIDATION PASS
```

- [ ] **Step 7: Commit implementation and validation sources**

```powershell
git add docs/superpowers/plans/2026-07-29-math6185-eight-slide-presentation.md work/math6185_eight_slide_package
git commit -m "feat: build eight-slide MATH6185 presentation package"
```

The final binary deliverables remain in the existing output folder and are delivered without replacing the source artifacts.
