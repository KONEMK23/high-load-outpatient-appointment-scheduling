# MATH6185 Obsidian Project Retrospective Design

## Objective

Convert the completed MATH6185 Case Study 1 project into a reusable Chinese-language Obsidian knowledge package. The notes should preserve the development lessons behind the work rather than reproduce the submitted report, model files, speaker script, or assessment answers verbatim.

## Audience and future use

The primary reader is the student returning to the vault before a future simulation, coursework, analytical-reporting, or technical-delivery project. The notes must make it easy to answer:

- How should a brief be converted into an executable project plan?
- How should related AnyLogic policies be structured and compared?
- How should warm-up, stopping rules, replications, confidence intervals, and parameter searches be justified?
- How should static validation be distinguished from a new simulation run?
- What Windows, PDF, XML, Word, PowerPoint, and Obsidian failures are likely to recur?
- What should be checked before submission and before an oral presentation?

## Vault location

Create the following folder and notes:

```text
Projects/
└─ MATH6185 Case Study 1/
   ├─ 00 - 项目复盘总览.md
   ├─ 01 - 需求解析与任务拆解.md
   ├─ 02 - AnyLogic模型架构与策略设计.md
   ├─ 03 - 实验设计与参数筛选.md
   ├─ 04 - 验证测试与质量门槛.md
   ├─ 05 - 报告提交与格式合规.md
   └─ 06 - PPT答辩与沟通准备.md
```

No existing vault note will be overwritten or deleted.

## Shared metadata

Every note will use YAML frontmatter with these fields:

```yaml
---
type: project-retrospective
project: MATH6185 Case Study 1
status: complete
created: 2026-07-29
course: MATH6185
tags:
  - project/math6185
---
```

Each specialist note will add one or more relevant hierarchical tags:

- `requirements/coursework`
- `simulation/anylogic`
- `simulation/experiment-design`
- `practice/validation`
- `delivery/report`
- `delivery/presentation`

## Navigation design

`00 - 项目复盘总览.md` is the map of content. It will:

- state the project objective and final outcome;
- summarise the Q2, Q3, and Q4 policy results;
- show the end-to-end workflow from brief to submission and presentation;
- link to all six specialist notes using Obsidian wikilinks;
- provide a compact reusable checklist for the next project.

Every specialist note will link back to the overview at the top and include contextual links to the other specialist notes where one stage depends on another. Links must use the exact note titles so that no unresolved wikilinks remain.

## Note contents

### 00 - 项目复盘总览

- Project question, constraints, and deliverables.
- Final Q2-Q4 comparison and conditional Q3 recommendation.
- End-to-end workflow.
- Seven-note navigation index.
- Highest-value reusable lessons.
- Next-project checklist.

### 01 - 需求解析与任务拆解

- Read the brief before building.
- Separate hard requirements, recommendations, and inferred choices.
- Convert questions Q1-Q4 into evidence and artifact requirements.
- Treat a strict maximum word count as a ceiling rather than a target or minimum.
- Track AI/peer acknowledgement separately from recommended report structure.
- Establish naming, output paths, and completion criteria at the start.

### 02 - AnyLogic模型架构与策略设计

- Shared model assumptions and the common comparison framework.
- Q2 first-available scheduling as the neutral benchmark.
- Q3 workload-responsive postponement:
  `round(alpha1 * N_U + alpha2 * N_U * N_U)`.
- Q4 hold-release-rebook logic, OCR selection, slot reassignment, and
  `DelayR.reduceDelay`.
- Why related models should differ only at the policy decision.
- Implementation mistakes to avoid and structural checks to retain.

### 03 - 实验设计与参数筛选

- Nominal-load calculation and implications of a 0.98 load.
- Warm-up of 1,000 served patients.
- Stopping rule of 50,000 arrivals.
- 250 independent replications and 95% confidence intervals.
- Relative confidence-interval half-width target.
- Pilot screening versus formal validation.
- Q3 and Q4 parameter-selection logic.
- Why feasible and optimal are not synonyms.

### 04 - 验证测试与质量门槛

- Evidence ladder: source requirements, model structure, fresh run output, report numbers, rendered artifacts, and final package.
- Difference between a new simulation run and static validation.
- `.alp` validation from raw UTF-8 bytes with `lxml`.
- PDF page, text, citation, metadata, font, and render checks.
- Cross-artifact numerical consistency.
- Hashes, filename checks, placeholder scans, and delivery-folder audits.
- Failed approaches and their fixes:
  Windows GBK stdout, PowerShell XML parsing, runtime Poppler paths,
  LibreOffice absence, and locked Office files.

### 05 - 报告提交与格式合规

- Answer-first analytical report structure.
- Linking literature to mechanisms rather than using references decoratively.
- Reporting means, confidence intervals, feasibility, trade-offs, and limitations.
- Word-count interpretation and concise writing.
- Reference and DOI verification.
- AI acknowledgement and author-metadata checks.
- Final four-file submission package and naming discipline.

### 06 - PPT答辩与沟通准备

- Convert the report into an eight-slide decision narrative.
- Explain the job of Slides 1-8.
- Keep backup evidence separate when unnecessary backup slides may broaden questions.
- Build a bilingual speaker script and a PPT-focused bilingual Q&A guide.
- Use listening keywords, one-number cues, and mechanism cues.
- Answer questions with: conclusion, one piece of evidence, then limitation.
- Audio rehearsal, timing, and pronunciation.
- Avoid live AnyLogic execution unless explicitly required.

## Content rules

- Write in clear Chinese; retain model names, variables, formulas, commands, and file extensions in English.
- Use short paragraphs, bullets, tables only for genuine comparisons, and fenced code blocks for commands.
- Preserve exact validated values where they are useful:
  Q2 overall 4.148 days; Q3 urgent 2.123, routine 9.053, overall 4.703;
  Q4 urgent 4.009, routine 4.574, overall 4.219, utilisation 97.911%.
- State that Q3 is recommended for the stated objective but needs a routine-patient guardrail.
- Do not state that the final audit was a fresh 250-replication run when it was a static validation of existing outputs and model structure.
- Do not copy long report passages, the full bibliography, the full presentation script, or the full Q&A guide.
- Local paths may be recorded in code blocks as provenance, but the notes must remain understandable if those paths later change.

## Source hierarchy

Use sources in this order:

1. The assignment brief `D:\Users\25470\Downloads\b6.pdf`.
2. Final submission artifacts under
   `outputs\zw1f25_MATH6185_submission`.
3. Validated Q2-Q4 run outputs and project scripts.
4. Final report and presentation artifacts.
5. Git design/implementation plans and recorded failure notes.

If sources conflict, prefer the assignment brief for requirements and the final validated submission package for reported results.

## Error handling and safety

- Read before writing when a target path already exists.
- Whole-file creation must fail rather than overwrite an unexpected existing note.
- Do not delete `欢迎.md`, daily notes, or MCP verification notes.
- After creation, read every note back and verify headings, frontmatter, tags, and wikilinks.
- If the Obsidian connection fails, restart or reconnect the app before retrying; do not write directly into an unknown vault filesystem path.

## Acceptance criteria

- Exactly seven new notes exist under `Projects/MATH6185 Case Study 1/`.
- The overview links to all six specialist notes and each specialist note links back.
- All notes contain valid shared frontmatter and the intended hierarchical tags.
- No unresolved internal link target is introduced.
- Q2-Q4 numerical values match the final validated outputs.
- The notes contain the major failures, fixes, and reusable checklists.
- Static validation and fresh simulation runs are explicitly distinguished.
- Existing vault notes remain unchanged.
