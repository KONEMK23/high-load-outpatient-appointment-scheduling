# MATH6185 Obsidian Project Retrospective Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create a seven-note Chinese Obsidian knowledge package that turns the completed MATH6185 Case Study 1 into reusable simulation-development, validation, reporting, and presentation experience.

**Architecture:** One overview note is the map of content and links to six focused notes. Each focused note owns one stage of the workflow, links back to the overview, records decisions and failure lessons, and ends with a reusable checklist. Creation uses the Obsidian MCP write interface; verification uses MCP readback, note listing, and tag listing so existing vault content is not modified.

**Tech Stack:** Obsidian Markdown, YAML frontmatter, Obsidian wikilinks, hierarchical tags, Obsidian Local REST API through the configured MCP bridge.

## Global Constraints

- Create exactly seven notes under `Projects/MATH6185 Case Study 1/`.
- Do not overwrite, patch, rename, or delete any existing vault note.
- Write explanations in Chinese while preserving English model names, variables, formulas, commands, file extensions, and technical terms where they aid precision.
- Every note must use `type: project-retrospective`, `project: MATH6185 Case Study 1`, `status: complete`, `created: 2026-07-29`, `course: MATH6185`, and `project/math6185`.
- Use the exact validated values: Q2 overall `4.148` days; Q3 urgent `2.123`, routine `9.053`, overall `4.703`; Q4 urgent `4.009`, routine `4.574`, overall `4.219`, utilisation `97.911%`.
- State that Q3 is the conditional recommendation for the stated objective and requires a routine-patient guardrail.
- Explicitly distinguish a fresh 250-replication simulation run from static validation of existing outputs and model structure.
- Do not reproduce the full report, bibliography, presentation script, or Q&A guide.
- Preserve these pre-existing vault notes unchanged: `2026-07-22.md`, `Codex MCP write verification 2026-07-22.md`, and `欢迎.md`.
- If a target note already exists at preflight, stop creation and report the collision instead of replacing it.

## File Structure

- Create: `Projects/MATH6185 Case Study 1/00 - 项目复盘总览.md` — project map, results summary, workflow, navigation, and next-project checklist.
- Create: `Projects/MATH6185 Case Study 1/01 - 需求解析与任务拆解.md` — brief interpretation and deliverable planning.
- Create: `Projects/MATH6185 Case Study 1/02 - AnyLogic模型架构与策略设计.md` — common model architecture and Q2-Q4 policy mechanisms.
- Create: `Projects/MATH6185 Case Study 1/03 - 实验设计与参数筛选.md` — load, warm-up, replications, confidence intervals, and search logic.
- Create: `Projects/MATH6185 Case Study 1/04 - 验证测试与质量门槛.md` — evidence ladder, static checks, run checks, and recurring Windows failures.
- Create: `Projects/MATH6185 Case Study 1/05 - 报告提交与格式合规.md` — report narrative, references, formatting, metadata, naming, and submission audit.
- Create: `Projects/MATH6185 Case Study 1/06 - PPT答辩与沟通准备.md` — eight-slide narrative, bilingual rehearsal, Q&A method, and delivery checklist.
- Modify: none.

---

### Task 1: Vault Safety Preflight

**Files:**
- Inspect: all notes under the current Obsidian vault.
- Create: none.
- Modify: none.

**Interfaces:**
- Consumes: the active Obsidian MCP connection and the seven exact target paths in the file structure.
- Produces: confirmation that all seven target paths are absent and a baseline snapshot of existing note paths and tags.

- [ ] **Step 1: List the current vault recursively**

Call `obsidian_list_notes` with `folder=""` and `recursive=true`.

Expected baseline includes:

```text
2026-07-22.md
Codex MCP write verification 2026-07-22.md
欢迎.md
```

- [ ] **Step 2: Check every target path for collisions**

Compare the returned paths against these exact targets:

```text
Projects/MATH6185 Case Study 1/00 - 项目复盘总览.md
Projects/MATH6185 Case Study 1/01 - 需求解析与任务拆解.md
Projects/MATH6185 Case Study 1/02 - AnyLogic模型架构与策略设计.md
Projects/MATH6185 Case Study 1/03 - 实验设计与参数筛选.md
Projects/MATH6185 Case Study 1/04 - 验证测试与质量门槛.md
Projects/MATH6185 Case Study 1/05 - 报告提交与格式合规.md
Projects/MATH6185 Case Study 1/06 - PPT答辩与沟通准备.md
```

Expected: none of the seven paths exists. If any exists, stop before writing.

- [ ] **Step 3: Record the current tag baseline**

Call `obsidian_list_tags`.

Expected: the call succeeds; the exact baseline may be empty. This snapshot will later prove which project tags were added.

- [ ] **Step 4: Verify the Obsidian connection remains writable without changing content**

Use only the successful list/read calls as the connection check. Do not create a temporary note, because the acceptance criterion is exactly seven new notes.

Expected: both note and tag listings return successfully.

### Task 2: Create the Overview Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/00 - 项目复盘总览.md`
- Modify: none.

**Interfaces:**
- Consumes: exact Q2-Q4 values and the six specialist note titles.
- Produces: the navigation hub used by Tasks 3-8.

- [ ] **Step 1: Write the overview note**

Call `obsidian_write_note` with `mode="overwrite"` only after Task 1 proves the path is absent. The content must contain this frontmatter:

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

Use these headings:

```markdown
# MATH6185 Case Study 1：项目复盘总览
## 项目一句话概括
## 最终成果
## 三个策略的结果对照
## 为什么推荐 Q3
## 从任务书到提交的完整流程
## 专题笔记导航
## 最值得复用的经验
## 下一个项目的启动清单
```

The results table must include all four global-constraint values, and “为什么推荐 Q3” must state that its lower urgent waiting time supports the stated objective while `9.053` routine days requires a guardrail. The navigation section must contain exact wikilinks to all six specialist note titles.

- [ ] **Step 2: Read the overview back**

Call `obsidian_get_note` for the exact overview path.

Expected: frontmatter contains the six shared fields; body contains all nine headings, all six specialist wikilinks, `4.148`, `2.123`, `9.053`, `4.703`, `4.009`, `4.574`, `4.219`, and `97.911%`.

- [ ] **Step 3: Verify overview scope**

Confirm the readback does not contain the complete report bibliography, the full presentation script, or a copied full Q&A list.

Expected: the note is a concise project map rather than an archive dump.

### Task 3: Create the Requirements Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/01 - 需求解析与任务拆解.md`
- Modify: none.

**Interfaces:**
- Consumes: assignment-brief lessons and the overview title.
- Produces: a reusable brief-to-plan workflow.

- [ ] **Step 1: Write the requirements note**

Use shared frontmatter and add `requirements/coursework`. Start with:

```markdown
[[00 - 项目复盘总览|返回项目总览]]
```

Use these headings:

```markdown
# 需求解析与任务拆解
## 先分清三类要求
## 把 Q1-Q4 转换成证据和文件
## 7500 字限制应该如何理解
## 文件命名和身份信息
## 完成标准要在开工时定义
## 本项目踩过的坑
## 可复用检查清单
```

Explain that a strict maximum is a ceiling, not a minimum or target; separate hard requirements from recommendations and inferred choices; track AI/peer acknowledgement independently; and define naming, output paths, model evidence, report evidence, and completion checks before building.

- [ ] **Step 2: Read the requirements note back**

Expected: the back-link resolves by exact title, both tags are present, all seven headings exist, and the note includes the phrases “硬性要求”, “建议”, “推断选择”, “上限”, and “完成标准”.

### Task 4: Create the AnyLogic Architecture Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/02 - AnyLogic模型架构与策略设计.md`
- Modify: none.

**Interfaces:**
- Consumes: the common model comparison framework and Q2-Q4 policy definitions.
- Produces: a reusable architecture pattern for related simulation policies.

- [ ] **Step 1: Write the architecture note**

Use shared frontmatter and add `simulation/anylogic`. Include the overview back-link and these headings:

```markdown
# AnyLogic模型架构与策略设计
## 共同模型骨架
## Q2：First Available Slot
## Q3：Workload-responsive Postponement
## Q4：Hold-Release-Rebook
## 如何保证策略比较公平
## 常见实现错误
## 可复用检查清单
```

Define the Q3 rule exactly as:

```text
round(alpha1 * N_U + alpha2 * N_U * N_U)
```

Explain Q4 hold/release/rebook flow, OCR patient selection, slot reassignment, and `DelayR.reduceDelay`. State that related models should share arrivals, capacity, stopping rule, metrics, and randomisation approach and differ only at the policy decision.

- [ ] **Step 2: Read the architecture note back**

Expected: shared metadata and both tags are present; all seven headings exist; the exact formula, `DelayR.reduceDelay`, Q2, Q3, Q4, OCR, and the overview back-link appear.

### Task 5: Create the Experiment Design Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/03 - 实验设计与参数筛选.md`
- Modify: none.

**Interfaces:**
- Consumes: final experimental settings and policy-search reasoning.
- Produces: a reusable simulation experiment template.

- [ ] **Step 1: Write the experiment note**

Use shared frontmatter and add `simulation/experiment-design`. Include the overview back-link and these headings:

```markdown
# 实验设计与参数筛选
## 先算名义负载
## Warm-up、终止条件与独立重复
## 置信区间与精度门槛
## Pilot Screening 与 Formal Validation
## Q3 参数选择逻辑
## Q4 参数选择逻辑
## Feasible 不等于 Optimal
## 可复用检查清单
```

Record nominal load `0.98`, warm-up `1,000 served patients`, stopping rule `50,000 arrivals`, `250 independent replications`, `95% confidence intervals`, and relative confidence-interval half-width. Explain why pilot screening may narrow candidates but does not replace formal validation.

- [ ] **Step 2: Read the experiment note back**

Expected: both tags and the overview back-link are present; all eight headings exist; `0.98`, `1,000`, `50,000`, `250`, `95%`, “Pilot Screening”, “Formal Validation”, “Feasible”, and “Optimal” appear.

### Task 6: Create the Verification Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/04 - 验证测试与质量门槛.md`
- Modify: none.

**Interfaces:**
- Consumes: validated workflows, tooling failures, and the fresh-run/static-check distinction.
- Produces: a reusable evidence ladder and final-quality gate.

- [ ] **Step 1: Write the verification note**

Use shared frontmatter and add `practice/validation`. Include the overview back-link and these headings:

```markdown
# 验证测试与质量门槛
## 证据阶梯
## 新仿真运行与静态验证的边界
## AnyLogic `.alp` 结构检查
## PDF、Word 与 PPT 渲染检查
## 跨文件数值一致性
## 最终交付文件夹审计
## 本项目失败记录与修复
## 可复用检查清单
```

The evidence ladder must run from source requirements through model structure, fresh run output, report numbers, rendered artifacts, and final package. Record these exact lessons:

```text
Windows GBK stdout -> set PYTHONIOENCODING=utf-8
PowerShell [xml] parsing -> validate raw UTF-8 bytes with Python lxml
Poppler not on PATH -> use the bundled/runtime absolute path
LibreOffice unavailable -> use available Office/PDF render tooling
Office file locked -> close the application or write to a new output path
```

State unambiguously that the final model audit was static validation of existing outputs and model structure, not a newly executed 250-replication run.

- [ ] **Step 2: Read the verification note back**

Expected: both tags and the overview back-link are present; all eight headings exist; `PYTHONIOENCODING=utf-8`, `lxml`, `Poppler`, `LibreOffice`, “locked”, “静态验证”, “新仿真运行”, and `250` appear.

### Task 7: Create the Report and Submission Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/05 - 报告提交与格式合规.md`
- Modify: none.

**Interfaces:**
- Consumes: final report, reference-verification, format, metadata, and submission lessons.
- Produces: a reusable report-to-delivery checklist.

- [ ] **Step 1: Write the report note**

Use shared frontmatter and add `delivery/report`. Include the overview back-link and these headings:

```markdown
# 报告提交与格式合规
## Answer-first 报告结构
## 如何把文献连接到模型机制
## 结果应该怎样呈现
## 字数与精简
## 参考文献真实性检查
## AI 声明、作者信息与隐藏元数据
## 四文件提交包
## 可复用检查清单
```

Explain that results reporting includes means, confidence intervals, feasibility, trade-offs, and limitations; references should support mechanisms rather than decorate prose; DOI/title/author/year should be checked; visible identity and PDF/Word metadata should be audited separately. Name the four deliverable types as one PDF report plus the Q2, Q3, and Q4 `.alp` models without copying the full report.

- [ ] **Step 2: Read the report note back**

Expected: both tags and the overview back-link are present; all eight headings exist; “confidence interval”, “DOI”, “隐藏元数据”, “PDF”, `.alp`, Q2, Q3, and Q4 appear.

### Task 8: Create the Presentation Note

**Files:**
- Create: `Projects/MATH6185 Case Study 1/06 - PPT答辩与沟通准备.md`
- Modify: none.

**Interfaces:**
- Consumes: the eight-slide deck structure, bilingual rehearsal process, and PPT-focused question strategy.
- Produces: a reusable presentation and oral-Q&A playbook.

- [ ] **Step 1: Write the presentation note**

Use shared frontmatter and add `delivery/presentation`. Include the overview back-link and these headings:

```markdown
# PPT答辩与沟通准备
## 八页主线分别承担什么任务
## 为什么备用页要谨慎
## 中英双语演讲稿如何使用
## 如何听懂问题
## 如何在没有时间检索时回答
## 回答结构：结论、证据、限制
## 彩排、计时与发音
## 可复用检查清单
```

Summarise Slides 1-8 as context, system, Q2 baseline, Q3 mechanism, Q3 evidence, Q4 mechanism/evidence, comparison/recommendation, and conclusion/limitations. Recommend listening for policy names, numbers, mechanism terms, and constraint words. Explain the three-part answer pattern and state that live AnyLogic execution should not be volunteered unless explicitly required.

- [ ] **Step 2: Read the presentation note back**

Expected: both tags and the overview back-link are present; all eight headings exist; “Slides 1-8”, “结论”, “证据”, “限制”, “AnyLogic”, “计时”, and “发音” appear.

### Task 9: Full Vault Acceptance Audit

**Files:**
- Inspect: all seven new project notes.
- Verify unchanged: `2026-07-22.md`, `Codex MCP write verification 2026-07-22.md`, and `欢迎.md`.
- Modify: only a newly created project note if its own readback fails an acceptance check.

**Interfaces:**
- Consumes: the seven notes produced by Tasks 2-8 and the Task 1 baseline.
- Produces: evidence that the project package satisfies the design specification.

- [ ] **Step 1: List the project folder recursively**

Call `obsidian_list_notes` with `folder="Projects/MATH6185 Case Study 1"` and `recursive=true`.

Expected: exactly these seven `.md` paths and no eighth temporary or duplicate note.

- [ ] **Step 2: Read all seven notes**

Call `obsidian_get_note` for each exact path.

Expected for every note:

```text
type: project-retrospective
project: MATH6185 Case Study 1
status: complete
created: 2026-07-29
course: MATH6185
project/math6185
```

- [ ] **Step 3: Audit internal links**

Verify that the overview contains six exact specialist wikilinks and every specialist contains:

```markdown
[[00 - 项目复盘总览|返回项目总览]]
```

Expected: each referenced title corresponds to one of the seven listed notes; no unresolved project-note target is introduced.

- [ ] **Step 4: Audit tags**

Call `obsidian_list_tags`.

Expected tags include:

```text
project/math6185
requirements/coursework
simulation/anylogic
simulation/experiment-design
practice/validation
delivery/report
delivery/presentation
```

- [ ] **Step 5: Audit numerical and methodological consistency**

Search the seven readbacks for the canonical values and claims.

Expected:

```text
Q2 overall = 4.148 days
Q3 urgent = 2.123 days
Q3 routine = 9.053 days
Q3 overall = 4.703 days
Q4 urgent = 4.009 days
Q4 routine = 4.574 days
Q4 overall = 4.219 days
Q4 utilisation = 97.911%
Q3 recommendation = conditional, with routine-patient guardrail
final audit = static validation, not a fresh 250-replication run
```

- [ ] **Step 6: Verify pre-existing notes remain present**

List the vault root recursively and confirm the three baseline notes still exist. Do not compare or alter their content unless an unexpected missing path is found.

Expected: all three original notes remain present and the only new paths are the seven project notes.

- [ ] **Step 7: Report completion**

Return links or exact vault-relative paths for all seven notes, summarize the completed checks, and disclose any check that could not be performed. Do not claim success for a failed or skipped check.
