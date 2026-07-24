# MATH6185 Oral Presentation Design

## Purpose

Prepare an adaptable oral-presentation package for the MATH6185 Case Study 1 assessment. The presentation must demonstrate the submitter's understanding of the project, engagement with relevant literature, technically sound methodology and results, and clear professional communication during both the talk and examiner questions.

## Confirmed Constraints

- The presentation is an individual, in-person oral assessment followed by examiner discussion and questions.
- The oral presentation is worth 30% of the module assessment and will be marked independently by two markers.
- The brief does not specify the talk duration or slide count.
- The deck will therefore support an 8–12 minute talk.
- Slides and speaker script will be in English.
- The Q&A guide will provide English answers with Chinese explanations.
- There will be no live AnyLogic run. Model structure and verified results will be shown through diagrams, charts, tables, and backup technical slides.
- The report and the three submitted `.alp` models will not be modified.

## Narrative Approach

Use a result-first defence narrative. Lead with the decision and trade-off, then show the modelling and statistical evidence that supports it.

Central message:

> The workload-responsive Q3 policy produces the largest reduction in urgent waiting time while satisfying the five-day overall-mean constraint, but it does so at a substantial cost to routine patients. The Q4 OCR policy is operationally gentler but delivers only a modest urgent-wait improvement.

## Core Slide Structure

1. **Title**
   - Simulation-Based Appointment Scheduling for Urgent and Routine Healthcare Demand
   - MATH6185, Case Study 1, Submitter zw1f25

2. **Why this problem is difficult**
   - Daily capacity of six slots against total arrival rate 5.88 per day
   - Approximately 98% utilisation before policy effects
   - Conflict between urgent access, routine delay, utilisation, and the overall five-day constraint
   - Short literature framing: FAS, capacity reservation, and dynamic priority

3. **Answer first: Q3 gives the largest urgent-wait reduction**
   - Comparison chart for FAS, Q3, and Q4
   - Urgent, routine, and overall means
   - Immediate statement of the Q3 benefit and fairness cost

4. **Model and experimental design**
   - Simplified AnyLogic process diagram
   - Independent Poisson arrivals at 3.69 and 2.19 per day
   - Six slots per working day
   - Warm-up after 1,000 served patients
   - Stop at 50,000 arrivals
   - 250 independent replications and 95% confidence intervals from replication means

5. **Q2 validates the model**
   - FAS overall mean 4.148 days, 95% CI [3.978, 4.318]
   - Urgent and routine means both approximately 4.148 days
   - Warm-up sensitivity across 500, 1,000, and 2,000 served patients
   - Explain why class equality is an internal validation check

6. **Q3: large urgent improvement, paid for by routine delay**
   - Policy rule: postpone a routine booking by `round(alpha1*N_U + alpha2*N_U^2)` slots relative to FAS
   - Selected parameters: alpha1 = 0, alpha2 = 0.35
   - Urgent 2.123 days, routine 9.053 days, overall 4.703 days
   - Overall 95% CI upper limit 4.893 days
   - Urgent reduction of 48.8% relative to FAS
   - Literature link to dynamic priority and workload-responsive scheduling

7. **Q4: controlled release preserves utilisation**
   - Selected parameters: eta = 0.2, L = 0, H = 3
   - Initial protection of the first H daily slots from routine bookings
   - End-of-day release and rebooking of eligible OCR patients
   - Urgent 4.009 days, routine 4.574 days, overall 4.219 days
   - Appointment-slot utilisation 97.911%
   - Average saving of 0.515 days per successful OCR rebooking
   - Literature link to Patrick and Puterman (2007)

8. **Recommendation and limitations**
   - Recommend Q3 for the stated urgent-wait objective
   - Explicitly identify the routine-wait fairness problem
   - Recommend a routine service-level constraint for real implementation
   - State modelling limitations and that finite parameter searches do not prove global optimality

## Backup Slides

- Full confidence intervals and relative half-widths
- Warm-up sensitivity table
- Q3 parameter screen for alpha2 = 0.35, 0.38, and 0.45
- Q4 policy screen and distinction between the event-equivalent runner and direct AnyLogic validation
- AnyLogic implementation map: `appts`, `N_U`, booking, rebooking, and statistics
- Key literature and how it influenced the policies
- AI-assistance disclosure and ownership statement

## Visual Design

- 16:9 widescreen format
- Clean academic style using navy, white, and teal
- Amber or red only for constraints, risks, or trade-offs
- No decorative stock imagery
- One main comparison chart, one simplified model-flow diagram, and one Q4 rebooking diagram
- One main message per slide
- Minimum font sizes consistent with the presentation skill requirements
- Short source citations in relevant slide footers
- English-only audience-facing copy

## Deliverables

1. `zw1f25_MATH6185_Presentation.pptx`
   - Eight core slides plus backup slides
   - English speaker notes embedded where supported

2. `zw1f25_MATH6185_Presentation.pdf`
   - Rendered presentation backup

3. `zw1f25_MATH6185_Speaker_Script.docx`
   - English script organised by slide
   - Optional sentences marked for omission in a shorter talk

4. `zw1f25_MATH6185_QA_Guide.docx`
   - Likely examiner questions
   - Concise English answers
   - Chinese explanations and study cues

## Evidence and Data Sources

- Teacher brief: `D:\Users\25470\Downloads\b6.pdf`
- Final report: `outputs\zw1f25_MATH6185_submission\zw1f25_MATH6185_Report.pdf`
- Submitted Q2, Q3, and Q4 AnyLogic models in the same submission directory
- User-provided successful 250-replication outputs for Q2, Q3, and Q4
- Publisher DOI pages for the nine references in the final report

## Quality Assurance

- Cross-check every displayed number against the final report and user-provided outputs.
- Verify that chart values, labels, and policy parameters agree across slides, notes, and Q&A material.
- Render every slide to PNG and inspect at full size.
- Run overflow and overlap checks and fix every unintended issue.
- Render both Word documents and inspect all pages for clipping, broken tables, or unreadable text.
- Confirm that the deck is intelligible without a live AnyLogic run.
- Confirm that no report or `.alp` file has been changed.

## Out of Scope

- Editing the submitted report
- Editing or rerunning the three AnyLogic models
- Live simulation during the oral presentation
- Claiming global optimality for the selected policies
