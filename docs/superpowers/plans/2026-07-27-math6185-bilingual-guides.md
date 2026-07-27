# MATH6185 Bilingual Presentation Guides Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create two bilingual MATH6185 study guides that preserve the source English content, add faithful Chinese translations, and connect every Q&A item to the relevant PowerPoint slide.

**Architecture:** A focused Python builder will copy each source DOCX, add two custom paragraph styles, and insert mapping/translation paragraphs at verified source anchors. A separate validation script will compare source and output content, count required additions, verify slide mappings, and confirm the DOCX packages open successfully.

**Tech Stack:** Bundled Python 3, `python-docx`, WordprocessingML XML insertion, `zipfile`, bundled document render helper when LibreOffice is available.

## Global Constraints

- Preserve the original DOCX and PPTX files.
- Do not alter source English wording, numerical results, parameters, conclusions, tables, or evidence lines.
- Create `zw1f25_MATH6185_QA_Guide_Bilingual.docx` and `zw1f25_MATH6185_Speaker_Script_Bilingual.docx` in the existing presentation output folder.
- Add a PowerPoint mapping and full Chinese answer translation for all 22 Q&A items.
- Add Chinese translations for the delivery guide, purpose, full suggested wording, and optional shorter-version instructions for Slides 1-8.
- Chinese text is a study aid and is not intended to be spoken during the English presentation.

---

### Task 1: Build deterministic bilingual document generator

**Files:**
- Create: `work/math6185_bilingual_guides/build_bilingual_guides.py`
- Test: `work/math6185_bilingual_guides/test_bilingual_guides.py`
- Read: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide.docx`
- Read: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script.docx`

**Interfaces:**
- Consumes: `Path` objects for the source Q&A DOCX, source speaker-script DOCX, and output directory.
- Produces: `build_qa_guide(source: Path, target: Path) -> None` and `build_speaker_script(source: Path, target: Path) -> None`.
- Produces constants `QA_SLIDE_MAP`, `QA_ANSWER_TRANSLATIONS`, `SCRIPT_EXACT_TRANSLATIONS`, and `SCRIPT_PREFIX_TRANSLATIONS`.

- [ ] **Step 1: Write failing tests for required coverage**

```python
def test_translation_and_mapping_coverage():
    assert set(QA_SLIDE_MAP) == set(range(1, 23))
    assert set(QA_ANSWER_TRANSLATIONS) == set(range(1, 23))
    assert set(QA_SLIDE_MAP.values()) == {
        "Slide 2", "Slides 5 and 13", "Slides 5 and 10", "Slide 10",
        "Slide 4", "Slides 4 and 9", "Slide 9", "Slides 6 and 11",
        "Slide 11", "Slides 8 and 11", "Slides 7 and 12", "Slide 12",
        "Slide 7", "Slides 2 and 14", "Slide 8", "Slide 13", "Slide 14",
    }

def test_source_documents_have_expected_anchors():
    qa = Document(QA_SOURCE)
    script = Document(SCRIPT_SOURCE)
    assert sum(p.text.startswith("Q") and ". " in p.text for p in qa.paragraphs) == 22
    assert sum(p.text.startswith("Answer in English:") for p in qa.paragraphs) == 22
    assert sum(p.text.startswith("Slide ") and " - " in p.text for p in script.paragraphs) == 8
    assert sum(p.text == "Suggested wording" for p in script.paragraphs) == 8
```

- [ ] **Step 2: Run tests and confirm they fail before the builder exists**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\test_bilingual_guides.py
```

Expected: import failure for `build_bilingual_guides`.

- [ ] **Step 3: Implement paragraph insertion and styles**

Use an XML-safe helper that inserts a new paragraph immediately after an existing paragraph:

```python
def insert_after(paragraph, text: str, style: str):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    inserted = Paragraph(new_p, paragraph._parent)
    inserted.style = style
    inserted.add_run(text)
    return inserted
```

Add two styles if absent:

```python
def ensure_styles(doc: Document) -> None:
    if "PPT Mapping" not in doc.styles:
        style = doc.styles.add_style("PPT Mapping", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Aptos"
        style.font.size = Pt(9)
        style.font.bold = True
        style.font.color.rgb = RGBColor(31, 111, 139)
    if "Chinese Translation" not in doc.styles:
        style = doc.styles.add_style("Chinese Translation", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Microsoft YaHei"
        style.font.size = Pt(9)
        style.font.color.rgb = RGBColor(55, 65, 81)
```

Insert `对应PPT：...` after each Q&A heading, and insert `中文对照：...` after every `Answer in English:` paragraph. Insert the speaker-script translations immediately after the exact English paragraph they translate. Save to new target paths without changing source files.

- [ ] **Step 4: Run unit tests**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\test_bilingual_guides.py
```

Expected: all coverage and source-anchor tests pass.

- [ ] **Step 5: Commit generator and tests**

```powershell
git add work/math6185_bilingual_guides/build_bilingual_guides.py work/math6185_bilingual_guides/test_bilingual_guides.py
git commit -m "feat: generate bilingual MATH6185 guides"
```

### Task 2: Generate outputs and verify content preservation

**Files:**
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide_Bilingual.docx`
- Create: `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Speaker_Script_Bilingual.docx`
- Create: `work/math6185_bilingual_guides/validate_bilingual_guides.py`

**Interfaces:**
- Consumes: the two source DOCX files and the two generated bilingual DOCX files.
- Produces: exit code `0` and a validation summary containing `QA mappings: 22`, `QA translations: 22`, `Speaker slides: 8`, and `English preservation: PASS`.

- [ ] **Step 1: Write failing output-validation checks**

```python
def assert_subsequence(source_items, output_items):
    pos = 0
    for item in output_items:
        if pos < len(source_items) and item == source_items[pos]:
            pos += 1
    assert pos == len(source_items)

assert len([p for p in qa_out.paragraphs if p.text.startswith("对应PPT：")]) == 22
assert len([p for p in qa_out.paragraphs if p.text.startswith("中文对照：")]) == 22
assert len([p for p in script_out.paragraphs if p.text.startswith("中文对照：")]) >= 28
assert_subsequence(source_qa_text, filtered_qa_output_text)
assert_subsequence(source_script_text, filtered_script_output_text)
```

- [ ] **Step 2: Run validation before output generation**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\validate_bilingual_guides.py
```

Expected: failure because bilingual output files do not yet exist.

- [ ] **Step 3: Generate the two bilingual DOCX files**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\build_bilingual_guides.py
```

Expected: both target DOCX paths are printed and exist with non-zero size.

- [ ] **Step 4: Run structural validation**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\validate_bilingual_guides.py
```

Expected:

```text
QA mappings: 22
QA translations: 22
Speaker slides: 8
English preservation: PASS
DOCX package integrity: PASS
```

- [ ] **Step 5: Inspect representative extracted sections**

Print Q4, Q9, Q13, Q18, and Q22 from the bilingual Q&A guide, plus Slides 1, 4, 6, and 8 from the bilingual script. Confirm each section follows the approved English -> Chinese ordering and uses the specified slide mapping.

### Task 3: Render or perform documented fallback QA

**Files:**
- Read: both generated bilingual DOCX files.
- Use: `D:/Users/25470/.codex/plugins/cache/openai-primary-runtime/documents/26.723.12215/skills/documents/render_docx.py`
- Create only as temporary QA artifacts: `.tmp_math6185_bilingual_render/`

**Interfaces:**
- Consumes: the two generated bilingual DOCX files.
- Produces: page PNGs for every output when LibreOffice is available, otherwise a recorded `LibreOffice unavailable` fallback and structural validation evidence.

- [ ] **Step 1: Attempt canonical DOCX rendering**

Run the bundled renderer separately for each output:

```powershell
& $PYTHON $RENDERER $QA_OUTPUT --output_dir .tmp_math6185_bilingual_render\qa
& $PYTHON $RENDERER $SCRIPT_OUTPUT --output_dir .tmp_math6185_bilingual_render\script
```

Expected when LibreOffice is present: `page-<N>.png` files for both documents.

- [ ] **Step 2: Inspect every rendered page**

Check for clipped Chinese glyphs, broken headings, mapping lines separated from their questions, excessive gaps, and disrupted tables. If a defect is found, adjust only the inserted styles/spacing, regenerate, and rerender.

- [ ] **Step 3: Apply the documented fallback if LibreOffice is unavailable**

If rendering fails specifically with `FileNotFoundError` for LibreOffice/`soffice`, retain the structurally validated DOCX files, record that visual render QA was unavailable, and do not claim visual-render success.

- [ ] **Step 4: Run final verification**

Run:

```powershell
& $PYTHON work\math6185_bilingual_guides\test_bilingual_guides.py
& $PYTHON work\math6185_bilingual_guides\validate_bilingual_guides.py
git diff --check
```

Expected: tests pass, validation reports all required counts and English preservation, and `git diff --check` has no output.
