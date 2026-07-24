import fs from "node:fs";
import assert from "node:assert/strict";

const data = JSON.parse(
  fs.readFileSync(new URL("./source_data.json", import.meta.url), "utf8"),
);
const content = JSON.parse(
  fs.readFileSync(new URL("./presentation_content.json", import.meta.url), "utf8"),
);

assert.equal(data.experiment.replications, 250);
assert.equal(data.experiment.stopAtArrivals, 50000);
assert.equal(data.experiment.warmupServed, 1000);
assert.equal(data.q3.alpha2, 0.35);
assert.equal(data.q4.eta, 0.2);
assert.equal(data.q4.L, 0);
assert.equal(data.q4.H, 3);
assert.ok(data.q3.overallCI[1] < data.system.overallMeanConstraintDays);
assert.equal(data.references.length, 9);
assert.ok(data.references.every((item) => item.doi && item.title));
assert.equal(content.slides.filter((slide) => slide.kind === "core").length, 8);
assert.equal(content.slides.filter((slide) => slide.kind === "backup").length, 6);
assert.ok(content.qa.length >= 18);
assert.ok(
  content.slides.every(
    (slide) =>
      slide.id &&
      slide.title &&
      slide.purpose &&
      slide.visibleCopy &&
      slide.speakerNotes &&
      Array.isArray(slide.shortVersionOmissions) &&
      Array.isArray(slide.sources),
  ),
);
assert.ok(
  content.qa.every(
    (entry) =>
      entry.question &&
      entry.answerEnglish &&
      entry.explanationChinese &&
      entry.evidence,
  ),
);

console.log("SOURCE VALIDATION PASS");
