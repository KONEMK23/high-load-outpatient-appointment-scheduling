# MATH6185 Eight-Slide Presentation and Focused Q&A Design

## Objective

Create a presentation package whose visible scope is limited to the eight spoken slides, together with a bilingual Q&A guide that prepares only for questions arising directly from those slides and their immediate supporting logic.

## Source files

- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pptx`
- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation.pdf`
- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide_Bilingual.docx`

The source files will remain unchanged.

## Output files

- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pptx`
- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_Presentation_8_Slides.pdf`
- `outputs/zw1f25_MATH6185_presentation/zw1f25_MATH6185_QA_Guide_PPT_Focused_Bilingual.docx`

## Eight-slide presentation

The new deck will contain exact copies of source Slides 1-8 in the same order. Visible content, theme, masters, layouts, fonts, charts, notes, citations, footers, and slide numbering will be preserved. Source Slides 9-14 and their relationships will not appear in the new deck.

The PDF will contain the same eight pages and no backup pages.

## Focused bilingual Q&A guide

The guide will begin with a one-page quick-location table containing:

- slide number;
- listening keywords;
- corresponding question numbers.

It will then contain 18 questions grouped under Slides 1-8:

### Slide 1: Decision and result

1. What is the central decision problem?
2. What is the main result, and why is Q3 recommended?

### Slide 2: Capacity pressure and literature framing

3. How is the nominal load of 0.98 calculated, and what does it imply?
4. How did the literature shape the central trade-off?

### Slide 3: Cross-policy comparison

5. How do Q2, Q3 and Q4 compare?
6. Why recommend Q3 despite its effect on routine patients?

### Slide 4: Experimental design

7. Why use the same experiment design for all three policies and stop at 50,000 arrivals?
8. Why use a warm-up of 1,000 served patients?
9. Why use 250 independent replications and 95% confidence intervals?

### Slide 5: Q2 benchmark

10. Why is FAS a useful benchmark?
11. How was the Q2 benchmark validated?

### Slide 6: Q3 workload-responsive policy

12. How does Q3 work?
13. Why was alpha2 = 0.35 selected?
14. Why does routine waiting rise to 9.053 days?

### Slide 7: Q4 OCR policy

15. How does the Q4 OCR policy work, and what do eta, L and H mean?
16. Why is Q4's urgent improvement small despite 97.911% utilisation?

### Slide 8: Recommendation, limitations and next work

17. Would Q3 be implemented unchanged?
18. What are the main limitations and the next useful experiment?

Each question will contain, in order:

1. English question;
2. Chinese question;
3. English answer;
4. faithful Chinese answer;
5. listening keywords;
6. one-number or one-mechanism memory cue.

Answers will be concise enough for a 20-40 second oral response. The focused guide will omit backup-only topics such as the detailed H = 2 versus H = 3 comparison, event-equivalent screening terminology, the full confidence-interval table, and the separate AI-disclosure backup slide.

## Validation

- Confirm the new PPTX contains exactly eight slides corresponding to source Slides 1-8.
- Confirm slide text, notes, master/layout references, and visible media for Slides 1-8 are preserved.
- Render and inspect all eight slides and confirm there are no visual differences from the first eight source slides.
- Confirm the PDF contains exactly eight pages and matches the rendered PPTX.
- Confirm the Q&A guide contains one quick-location table, eight slide sections, and exactly 18 bilingual questions.
- Confirm every numerical claim in the focused guide matches the existing deck and validated simulation results.
- Preserve all original files.
