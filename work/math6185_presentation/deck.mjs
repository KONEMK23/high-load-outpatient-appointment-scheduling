import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  Presentation,
  PresentationFile,
} from "./tmp/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "..", "..");
const TMP = path.join(HERE, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const OUTPUT = path.join(
  ROOT,
  "outputs",
  "zw1f25_MATH6185_presentation",
  "zw1f25_MATH6185_Presentation.pptx",
);

const data = JSON.parse(
  await fs.readFile(path.join(HERE, "source_data.json"), "utf8"),
);
const content = JSON.parse(
  await fs.readFile(path.join(HERE, "presentation_content.json"), "utf8"),
);

const COLORS = {
  navy: "#17324D",
  teal: "#007F82",
  darkTeal: "#006164",
  paleTeal: "#DCEEEF",
  white: "#FFFFFF",
  ink: "#18222C",
  grey: "#687783",
  lightGrey: "#EEF2F4",
  amber: "#D68B00",
  paleAmber: "#FFF4D6",
  red: "#B23A48",
  paleRed: "#FBEAEC",
  line: "#CDD7DC",
};

const FONT = "Aptos";
const FONT_DISPLAY = "Aptos Display";
const SLIDE_W = 1280;
const SLIDE_H = 720;
const X = 72;
const RIGHT = 1208;

function addShape(
  slide,
  {
    geometry = "rect",
    left,
    top,
    width,
    height,
    fill = "none",
    lineFill = "none",
    lineWidth = 0,
    radius,
    name,
  },
) {
  const shape = slide.shapes.add({
    geometry,
    name,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
  });
  if (radius !== undefined) shape.borderRadius = radius;
  return shape;
}

function addText(
  slide,
  text,
  {
    left,
    top,
    width,
    height,
    fontSize = 20,
    color = COLORS.ink,
    bold = false,
    align = "left",
    valign = "top",
    typeface = FONT,
    fill = "none",
    lineFill = "none",
    lineWidth = 0,
    radius,
    insets = { left: 0, right: 0, top: 0, bottom: 0 },
    name,
  },
) {
  const shape = addShape(slide, {
    geometry: "rect",
    left,
    top,
    width,
    height,
    fill,
    lineFill,
    lineWidth,
    radius,
    name,
  });
  shape.text = text;
  shape.text.fontSize = fontSize;
  shape.text.color = color;
  shape.text.bold = bold;
  shape.text.typeface = typeface;
  shape.text.alignment = align;
  shape.text.verticalAlignment = valign;
  shape.text.insets = insets;
  return shape;
}

function addTitle(slide, title, subtitle = "", backup = false) {
  addText(slide, title, {
    left: X,
    top: 46,
    width: backup ? 945 : 1080,
    height: 54,
    fontSize: 38,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.navy,
    valign: "middle",
    name: "slide-title",
  });
  if (subtitle) {
    addText(slide, subtitle, {
      left: X,
      top: 106,
      width: 1080,
      height: 34,
      fontSize: 18,
      color: COLORS.grey,
      name: "slide-subtitle",
    });
  }
  if (backup) {
    addText(slide, "BACKUP", {
      left: 1060,
      top: 52,
      width: 148,
      height: 34,
      fontSize: 16,
      bold: true,
      color: COLORS.white,
      align: "center",
      valign: "middle",
      fill: COLORS.navy,
      radius: 17,
      name: "backup-label",
    });
  }
  addShape(slide, {
    left: X,
    top: 122,
    width: 1136,
    height: 3,
    fill: COLORS.teal,
  });
}

function addFooter(slide, sourceText, slideNumber) {
  addShape(slide, {
    left: X,
    top: 674,
    width: 1136,
    height: 1,
    fill: COLORS.line,
  });
  addText(slide, sourceText, {
    left: X,
    top: 683,
    width: 1020,
    height: 20,
    fontSize: 13,
    color: COLORS.grey,
    name: "source-footer",
  });
  addText(slide, String(slideNumber).padStart(2, "0"), {
    left: 1150,
    top: 681,
    width: 58,
    height: 22,
    fontSize: 14,
    bold: true,
    color: COLORS.navy,
    align: "right",
    name: "slide-number",
  });
}

function addSpeakerNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
}

function getSlideContent(id) {
  const item = content.slides.find((slide) => slide.id === id);
  if (!item) throw new Error(`Missing content for ${id}`);
  return item;
}

function sourceLine(item) {
  return item.sources.join(" ");
}

function addMetric(slide, value, label, options) {
  const { left, top, width, color = COLORS.teal, labelColor = COLORS.grey } =
    options;
  addText(slide, value, {
    left,
    top,
    width,
    height: 72,
    fontSize: 48,
    bold: true,
    typeface: FONT_DISPLAY,
    color,
    valign: "middle",
  });
  addText(slide, label, {
    left,
    top: top + 72,
    width,
    height: 46,
    fontSize: 18,
    bold: true,
    color: labelColor,
  });
}

function addMeanComparisonChart(slide) {
  slide.charts.add("bar", {
    position: { left: 82, top: 190, width: 780, height: 400 },
    categories: ["Urgent", "Routine", "Overall"],
    series: [
      {
        name: "Q2 FAS",
        values: [
          data.q2.urgentMean,
          data.q2.routineMean,
          data.q2.overallMean,
        ],
        fill: COLORS.grey,
        valuesFormatCode: "0.000",
      },
      {
        name: "Q3 workload-responsive",
        values: [
          data.q3.urgentMean,
          data.q3.routineMean,
          data.q3.overallMean,
        ],
        fill: COLORS.teal,
        valuesFormatCode: "0.000",
      },
      {
        name: "Q4 on-call routine",
        values: [
          data.q4.urgentMean,
          data.q4.routineMean,
          data.q4.overallMean,
        ],
        fill: COLORS.navy,
        valuesFormatCode: "0.000",
      },
    ],
    barOptions: {
      direction: "column",
      grouping: "clustered",
      gapWidth: 55,
    },
    hasLegend: true,
    legend: {
      position: "bottom",
      overlay: false,
      textStyle: { fill: COLORS.grey, fontSize: 16 },
    },
    xAxis: {
      textStyle: { fill: COLORS.ink, fontSize: 17, bold: true },
      line: { style: "solid", fill: COLORS.line, width: 1 },
      majorGridlines: null,
    },
    yAxis: {
      min: 0,
      max: 10,
      majorUnit: 2,
      title: {
        text: "Mean waiting time (days)",
        textStyle: { fill: COLORS.grey, fontSize: 16 },
      },
      textStyle: { fill: COLORS.grey, fontSize: 15 },
      line: { style: "solid", fill: COLORS.line, width: 1 },
      majorGridlines: { style: "solid", fill: COLORS.lightGrey, width: 1 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      textStyle: { fill: COLORS.ink, fontSize: 15, bold: true },
    },
    chartFill: COLORS.white,
    chartLine: { style: "solid", fill: "none", width: 0 },
    plotAreaFill: COLORS.white,
    plotAreaLine: { style: "solid", fill: "none", width: 0 },
  });
}

function addModelFlow(slide) {
  const labels = [
    ["Poisson arrivals", "Urgent 3.69/day\nRoutine 2.19/day"],
    ["Policy decision", "FAS | Q3 | Q4"],
    ["Appointment calendar", "6 slots/day"],
    ["Outcome statistics", "Waits, CIs, utilisation"],
  ];
  const boxes = [];
  const lefts = [78, 365, 652, 939];
  for (let i = 0; i < labels.length; i += 1) {
    const box = addShape(slide, {
      left: lefts[i],
      top: 188,
      width: 220,
      height: 134,
      fill: i === 1 ? COLORS.paleTeal : COLORS.white,
      lineFill: i === 1 ? COLORS.teal : COLORS.line,
      lineWidth: i === 1 ? 2 : 1,
      radius: 18,
      name: `flow-box-${i + 1}`,
    });
    boxes.push(box);
  }
  for (let i = 0; i < boxes.length - 1; i += 1) {
    slide.shapes.connect(boxes[i], boxes[i + 1], {
      kind: "straight",
      fromSide: "right",
      toSide: "left",
      line: { style: "solid", fill: COLORS.teal, width: 3 },
      tail: { type: "arrow", width: "med", length: "med" },
    });
  }
  for (let i = 0; i < labels.length; i += 1) {
    addText(slide, labels[i][0], {
      left: lefts[i] + 14,
      top: 206,
      width: 192,
      height: 34,
      fontSize: 21,
      bold: true,
      color: COLORS.navy,
      align: "center",
    });
    addText(slide, labels[i][1], {
      left: lefts[i] + 14,
      top: 250,
      width: 192,
      height: 54,
      fontSize: 17,
      color: COLORS.grey,
      align: "center",
      valign: "middle",
    });
  }
}

function addQ4RebookingFlow(slide) {
  const labels = [
    ["1", "Hold", "Protect H = 3 near-term positions"],
    ["2", "Release", "Scan unused slots at day end"],
    ["3", "Rebook", "Bring forward oldest eligible OCR patient"],
  ];
  const boxes = [];
  const lefts = [90, 405, 720];
  for (let i = 0; i < labels.length; i += 1) {
    const box = addShape(slide, {
      left: lefts[i],
      top: 180,
      width: 260,
      height: 150,
      fill: i === 2 ? COLORS.paleTeal : COLORS.white,
      lineFill: i === 2 ? COLORS.teal : COLORS.line,
      lineWidth: i === 2 ? 2 : 1,
      radius: 20,
      name: `q4-step-${i + 1}`,
    });
    boxes.push(box);
  }
  for (let i = 0; i < boxes.length - 1; i += 1) {
    slide.shapes.connect(boxes[i], boxes[i + 1], {
      kind: "straight",
      fromSide: "right",
      toSide: "left",
      line: { style: "solid", fill: COLORS.teal, width: 3 },
      tail: { type: "arrow", width: "med", length: "med" },
    });
  }
  for (let i = 0; i < labels.length; i += 1) {
    addText(slide, labels[i][0], {
      left: lefts[i] + 18,
      top: 198,
      width: 44,
      height: 44,
      fontSize: 22,
      bold: true,
      color: COLORS.white,
      align: "center",
      valign: "middle",
      fill: COLORS.teal,
      radius: 22,
    });
    addText(slide, labels[i][1], {
      left: lefts[i] + 72,
      top: 199,
      width: 164,
      height: 38,
      fontSize: 25,
      bold: true,
      color: COLORS.navy,
    });
    addText(slide, labels[i][2], {
      left: lefts[i] + 18,
      top: 254,
      width: 224,
      height: 60,
      fontSize: 17,
      color: COLORS.grey,
      align: "center",
      valign: "middle",
    });
  }
}

function addSimpleTable(
  slide,
  { rows, left, top, width, height, columnWidths, header = true, riskRows = [] },
) {
  const table = slide.tables.add({
    rows: rows.length,
    columns: rows[0].length,
    left,
    top,
    width,
    height,
    columnWidths,
    values: rows,
  });
  table.borders.assign({ style: "solid", fill: COLORS.line, width: 1 });
  table.cells
    .block({ row: 0, column: 0, rowCount: rows.length, columnCount: rows[0].length })
    .assign({
      textStyle: { fontSize: 16, color: COLORS.ink },
      margins: { left: 10, right: 10, top: 5, bottom: 5 },
      anchor: "middle",
    });
  if (header) {
    table.cells
      .block({ row: 0, column: 0, rowCount: 1, columnCount: rows[0].length })
      .assign({
        fill: COLORS.navy,
        textStyle: { fontSize: 17, color: COLORS.white, bold: true },
        anchor: "middle",
      });
  }
  for (let row = header ? 1 : 0; row < rows.length; row += 1) {
    if ((row - (header ? 1 : 0)) % 2 === 1) {
      table.cells
        .block({ row, column: 0, rowCount: 1, columnCount: rows[0].length })
        .assign({ fill: COLORS.lightGrey });
    }
  }
  for (const row of riskRows) {
    table.cells
      .block({ row, column: 0, rowCount: 1, columnCount: rows[0].length })
      .assign({ fill: COLORS.paleAmber });
  }
  return table;
}

function buildTitleSlide(presentation) {
  const item = getSlideContent("title");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.navy;
  addShape(slide, {
    left: 0,
    top: 0,
    width: 18,
    height: SLIDE_H,
    fill: COLORS.teal,
  });
  addText(slide, "MATH6185 | CASE STUDY 1", {
    left: 82,
    top: 62,
    width: 470,
    height: 32,
    fontSize: 17,
    bold: true,
    color: "#A9D7D8",
  });
  addText(slide, "Reducing urgent waits\nwithout overloading the clinic", {
    left: 82,
    top: 132,
    width: 730,
    height: 154,
    fontSize: 54,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.white,
  });
  addText(slide, item.visibleCopy.subtitle, {
    left: 84,
    top: 318,
    width: 720,
    height: 60,
    fontSize: 24,
    color: "#DCE7ED",
  });
  addShape(slide, {
    left: 882,
    top: 118,
    width: 254,
    height: 330,
    fill: COLORS.white,
    radius: 28,
  });
  addText(slide, "48.8%", {
    left: 910,
    top: 165,
    width: 198,
    height: 82,
    fontSize: 58,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.teal,
    align: "center",
  });
  addText(slide, "lower urgent mean\nunder Q3", {
    left: 910,
    top: 254,
    width: 198,
    height: 70,
    fontSize: 22,
    bold: true,
    color: COLORS.navy,
    align: "center",
  });
  addShape(slide, {
    left: 917,
    top: 346,
    width: 184,
    height: 2,
    fill: COLORS.line,
  });
  addText(slide, "Routine mean rises\nto 9.053 days", {
    left: 910,
    top: 370,
    width: 198,
    height: 56,
    fontSize: 17,
    color: COLORS.red,
    bold: true,
    align: "center",
  });
  addText(slide, item.visibleCopy.identity, {
    left: 84,
    top: 624,
    width: 700,
    height: 28,
    fontSize: 18,
    color: "#DCE7ED",
  });
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildProblemSlide(presentation) {
  const item = getSlideContent("problem");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title);
  addText(slide, "0.98", {
    left: 84,
    top: 178,
    width: 290,
    height: 112,
    fontSize: 82,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.teal,
  });
  addText(slide, "nominal load", {
    left: 88,
    top: 284,
    width: 280,
    height: 42,
    fontSize: 23,
    bold: true,
    color: COLORS.navy,
  });
  addText(slide, "Demand almost matches capacity, so every protection decision redistributes delay.", {
    left: 84,
    top: 354,
    width: 344,
    height: 116,
    fontSize: 25,
    bold: true,
    color: COLORS.ink,
  });
  addShape(slide, {
    left: 480,
    top: 176,
    width: 688,
    height: 228,
    fill: COLORS.lightGrey,
    radius: 22,
  });
  const facts = [
    ["Urgent demand", "3.69/day"],
    ["Routine demand", "2.19/day"],
    ["Available capacity", "6 slots/day"],
  ];
  facts.forEach((fact, index) => {
    const top = 204 + index * 62;
    addText(slide, fact[0], {
      left: 518,
      top,
      width: 300,
      height: 34,
      fontSize: 20,
      color: COLORS.grey,
      valign: "middle",
    });
    addText(slide, fact[1], {
      left: 860,
      top,
      width: 260,
      height: 34,
      fontSize: 25,
      bold: true,
      color: COLORS.navy,
      align: "right",
      valign: "middle",
    });
  });
  addText(slide, "Literature framing", {
    left: 480,
    top: 446,
    width: 270,
    height: 36,
    fontSize: 22,
    bold: true,
    color: COLORS.teal,
  });
  addText(slide, item.visibleCopy.literature, {
    left: 480,
    top: 490,
    width: 688,
    height: 92,
    fontSize: 20,
    color: COLORS.ink,
  });
  addText(slide, item.visibleCopy.question, {
    left: 480,
    top: 598,
    width: 688,
    height: 44,
    fontSize: 23,
    bold: true,
    color: COLORS.navy,
  });
  addFooter(slide, sourceLine(item), 2);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildComparisonSlide(presentation) {
  const item = getSlideContent("comparison");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title);
  addMeanComparisonChart(slide);
  addShape(slide, {
    left: 900,
    top: 190,
    width: 292,
    height: 226,
    fill: COLORS.paleTeal,
    radius: 22,
  });
  addText(slide, "Decision", {
    left: 928,
    top: 218,
    width: 234,
    height: 32,
    fontSize: 20,
    bold: true,
    color: COLORS.darkTeal,
  });
  addText(slide, "Choose Q3", {
    left: 928,
    top: 267,
    width: 234,
    height: 48,
    fontSize: 36,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.navy,
  });
  addText(slide, "Largest urgent reduction\nwhile overall mean remains feasible", {
    left: 928,
    top: 329,
    width: 234,
    height: 62,
    fontSize: 18,
    color: COLORS.ink,
  });
  addShape(slide, {
    left: 900,
    top: 442,
    width: 292,
    height: 144,
    fill: COLORS.paleAmber,
    radius: 22,
  });
  addText(slide, "Service risk", {
    left: 928,
    top: 466,
    width: 234,
    height: 30,
    fontSize: 20,
    bold: true,
    color: COLORS.amber,
  });
  addText(slide, "Routine mean rises to\n9.053 days", {
    left: 928,
    top: 510,
    width: 234,
    height: 58,
    fontSize: 23,
    bold: true,
    color: COLORS.navy,
  });
  addFooter(slide, sourceLine(item), 3);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildMethodSlide(presentation) {
  const item = getSlideContent("method");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.lightGrey;
  addTitle(slide, item.title);
  addModelFlow(slide);
  const metrics = [
    ["1,000", "served-patient warm-up"],
    ["50,000", "arrival stopping rule"],
    ["250", "independent replications"],
    ["95%", "confidence intervals"],
  ];
  metrics.forEach((metric, i) => {
    const left = 86 + i * 287;
    addText(slide, metric[0], {
      left,
      top: 390,
      width: 240,
      height: 66,
      fontSize: 42,
      bold: true,
      typeface: FONT_DISPLAY,
      color: i === 2 ? COLORS.teal : COLORS.navy,
      align: "center",
    });
    addText(slide, metric[1], {
      left,
      top: 462,
      width: 240,
      height: 48,
      fontSize: 17,
      bold: true,
      color: COLORS.grey,
      align: "center",
    });
  });
  addText(slide, "Precision target met: relative CI half-width <= 5% for every selected policy", {
    left: 214,
    top: 552,
    width: 852,
    height: 50,
    fontSize: 22,
    bold: true,
    color: COLORS.white,
    align: "center",
    valign: "middle",
    fill: COLORS.teal,
    radius: 25,
  });
  addFooter(slide, sourceLine(item), 4);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildQ2Slide(presentation) {
  const item = getSlideContent("q2");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title);
  addMetric(slide, "4.148", "overall mean waiting (days)", {
    left: 86,
    top: 184,
    width: 360,
  });
  addText(slide, "95% CI [3.978, 4.318]", {
    left: 88,
    top: 314,
    width: 360,
    height: 40,
    fontSize: 21,
    bold: true,
    color: COLORS.navy,
  });
  addText(slide, "Urgent 4.148   |   Routine 4.148", {
    left: 88,
    top: 370,
    width: 390,
    height: 40,
    fontSize: 20,
    color: COLORS.grey,
  });
  addShape(slide, {
    left: 520,
    top: 172,
    width: 650,
    height: 274,
    fill: COLORS.lightGrey,
    radius: 24,
  });
  addText(slide, "Warm-up sensitivity", {
    left: 556,
    top: 202,
    width: 560,
    height: 38,
    fontSize: 24,
    bold: true,
    color: COLORS.navy,
  });
  const warmups = [
    ["500 served", "4.133"],
    ["1,000 served", "4.148"],
    ["2,000 served", "4.172"],
  ];
  warmups.forEach((row, index) => {
    const top = 261 + index * 55;
    addText(slide, row[0], {
      left: 558,
      top,
      width: 250,
      height: 34,
      fontSize: 19,
      color: COLORS.grey,
      valign: "middle",
    });
    addText(slide, row[1], {
      left: 864,
      top,
      width: 220,
      height: 34,
      fontSize: 25,
      bold: true,
      color: COLORS.teal,
      align: "right",
      valign: "middle",
    });
  });
  addText(slide, "Strongly overlapping confidence intervals support a stable benchmark.", {
    left: 162,
    top: 510,
    width: 956,
    height: 74,
    fontSize: 25,
    bold: true,
    color: COLORS.navy,
    align: "center",
    valign: "middle",
    fill: COLORS.paleTeal,
    radius: 18,
  });
  addFooter(slide, sourceLine(item), 5);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildQ3Slide(presentation) {
  const item = getSlideContent("q3");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.navy;
  addText(slide, item.title, {
    left: X,
    top: 48,
    width: 1080,
    height: 58,
    fontSize: 38,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.white,
  });
  addShape(slide, {
    left: X,
    top: 122,
    width: 1136,
    height: 3,
    fill: COLORS.teal,
  });
  addText(slide, "Postponement = round(α1 NU + α2 NU²)", {
    left: 86,
    top: 174,
    width: 690,
    height: 62,
    fontSize: 31,
    bold: true,
    color: "#CDEBEC",
  });
  addText(slide, "α1 = 0   |   α2 = 0.35", {
    left: 88,
    top: 252,
    width: 540,
    height: 46,
    fontSize: 24,
    color: COLORS.white,
  });
  addText(slide, "Urgent backlog increases\nprotection nonlinearly.", {
    left: 88,
    top: 328,
    width: 520,
    height: 94,
    fontSize: 26,
    bold: true,
    color: COLORS.white,
  });
  addShape(slide, {
    left: 810,
    top: 164,
    width: 336,
    height: 246,
    fill: COLORS.white,
    radius: 24,
  });
  addText(slide, "2.123", {
    left: 846,
    top: 204,
    width: 264,
    height: 82,
    fontSize: 62,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.teal,
    align: "center",
  });
  addText(slide, "urgent mean (days)", {
    left: 846,
    top: 292,
    width: 264,
    height: 36,
    fontSize: 20,
    bold: true,
    color: COLORS.navy,
    align: "center",
  });
  addText(slide, "48.8% below FAS", {
    left: 846,
    top: 348,
    width: 264,
    height: 38,
    fontSize: 22,
    bold: true,
    color: COLORS.darkTeal,
    align: "center",
  });
  addShape(slide, {
    left: 86,
    top: 478,
    width: 1060,
    height: 112,
    fill: "#FFFFFF18",
    lineFill: "#FFFFFF30",
    lineWidth: 1,
    radius: 20,
  });
  addText(slide, "Routine", {
    left: 132,
    top: 503,
    width: 172,
    height: 28,
    fontSize: 17,
    color: "#BFD0DB",
  });
  addText(slide, "9.053 days", {
    left: 132,
    top: 535,
    width: 250,
    height: 38,
    fontSize: 27,
    bold: true,
    color: "#FFD79A",
  });
  addText(slide, "Overall", {
    left: 472,
    top: 503,
    width: 172,
    height: 28,
    fontSize: 17,
    color: "#BFD0DB",
  });
  addText(slide, "4.703 days", {
    left: 472,
    top: 535,
    width: 250,
    height: 38,
    fontSize: 27,
    bold: true,
    color: COLORS.white,
  });
  addText(slide, "Feasibility", {
    left: 812,
    top: 503,
    width: 172,
    height: 28,
    fontSize: 17,
    color: "#BFD0DB",
  });
  addText(slide, "Upper CI 4.893", {
    left: 812,
    top: 535,
    width: 250,
    height: 38,
    fontSize: 27,
    bold: true,
    color: COLORS.white,
  });
  addFooter(slide, sourceLine(item), 6);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildQ4Slide(presentation) {
  const item = getSlideContent("q4");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.lightGrey;
  addTitle(slide, item.title);
  addText(slide, "η = 0.2   |   L = 0   |   H = 3", {
    left: 842,
    top: 93,
    width: 330,
    height: 30,
    fontSize: 17,
    bold: true,
    color: COLORS.teal,
    align: "right",
  });
  addQ4RebookingFlow(slide);
  const metrics = [
    ["4.009", "urgent mean"],
    ["4.574", "routine mean"],
    ["4.219", "overall mean"],
    ["97.911%", "slot utilisation"],
  ];
  metrics.forEach((metric, index) => {
    const left = 76 + index * 286;
    addText(slide, metric[0], {
      left,
      top: 390,
      width: 242,
      height: 62,
      fontSize: 38,
      bold: true,
      typeface: FONT_DISPLAY,
      color: index === 0 ? COLORS.teal : COLORS.navy,
      align: "center",
    });
    addText(slide, metric[1], {
      left,
      top: 458,
      width: 242,
      height: 34,
      fontSize: 17,
      bold: true,
      color: COLORS.grey,
      align: "center",
    });
  });
  addText(slide, "0.515 days saved per rebooked OCR patient", {
    left: 236,
    top: 548,
    width: 808,
    height: 54,
    fontSize: 23,
    bold: true,
    color: COLORS.white,
    align: "center",
    valign: "middle",
    fill: COLORS.teal,
    radius: 27,
  });
  addFooter(slide, sourceLine(item), 7);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildRecommendationSlide(presentation) {
  const item = getSlideContent("recommendation");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title);
  addShape(slide, {
    left: 76,
    top: 170,
    width: 718,
    height: 176,
    fill: COLORS.navy,
    radius: 24,
  });
  addText(slide, "Choose Q3", {
    left: 112,
    top: 204,
    width: 300,
    height: 58,
    fontSize: 40,
    bold: true,
    typeface: FONT_DISPLAY,
    color: COLORS.white,
  });
  addText(slide, "Urgent mean 2.123 days\nOverall upper 95% CI 4.893 days", {
    left: 112,
    top: 270,
    width: 620,
    height: 58,
    fontSize: 20,
    color: "#DCE7ED",
  });
  addShape(slide, {
    left: 836,
    top: 170,
    width: 366,
    height: 176,
    fill: COLORS.paleAmber,
    radius: 24,
  });
  addText(slide, "Guardrail required", {
    left: 870,
    top: 206,
    width: 300,
    height: 38,
    fontSize: 24,
    bold: true,
    color: COLORS.amber,
  });
  addText(slide, "Routine mean = 9.053 days\nAdd a routine service limit.", {
    left: 870,
    top: 264,
    width: 300,
    height: 62,
    fontSize: 19,
    bold: true,
    color: COLORS.navy,
  });
  addText(slide, "What the evidence supports", {
    left: 82,
    top: 404,
    width: 350,
    height: 36,
    fontSize: 24,
    bold: true,
    color: COLORS.teal,
  });
  addText(slide, "Q3 is the best tested policy for the stated objective. It is not a free efficiency gain; it redistributes waiting from urgent to routine patients.", {
    left: 82,
    top: 456,
    width: 560,
    height: 116,
    fontSize: 22,
    color: COLORS.ink,
  });
  addText(slide, "Next experiment", {
    left: 704,
    top: 404,
    width: 350,
    height: 36,
    fontSize: 24,
    bold: true,
    color: COLORS.teal,
  });
  addText(slide, "Combine Q3's backlog signal with controlled OCR release, then test a routine service-level constraint.", {
    left: 704,
    top: 456,
    width: 464,
    height: 94,
    fontSize: 22,
    color: COLORS.ink,
  });
  addText(slide, "Limitations: stationary demand, deterministic service, no no-shows or overtime, finite parameter search.", {
    left: 82,
    top: 614,
    width: 1086,
    height: 32,
    fontSize: 16,
    color: COLORS.grey,
  });
  addFooter(slide, sourceLine(item), 8);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildCITableSlide(presentation) {
  const item = getSlideContent("ci_table");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title, "", true);
  addSimpleTable(slide, {
    rows: [
      ["Metric", "Mean", "95% confidence interval"],
      ...item.visibleCopy.rows,
    ],
    left: 172,
    top: 164,
    width: 936,
    height: 444,
    columnWidths: [360, 210, 366],
    riskRows: [3],
  });
  addFooter(slide, sourceLine(item), 9);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildWarmupSlide(presentation) {
  const item = getSlideContent("warmup");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.lightGrey;
  addTitle(slide, item.title, "", true);
  addSimpleTable(slide, {
    rows: [["Warm-up", "Overall mean", "95% confidence interval"], ...item.visibleCopy.rows],
    left: 198,
    top: 188,
    width: 884,
    height: 264,
    columnWidths: [270, 240, 374],
  });
  addText(slide, item.visibleCopy.takeaway, {
    left: 170,
    top: 506,
    width: 940,
    height: 80,
    fontSize: 24,
    bold: true,
    color: COLORS.navy,
    align: "center",
    valign: "middle",
    fill: COLORS.paleTeal,
    radius: 20,
  });
  addFooter(slide, sourceLine(item), 10);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildQ3ScreenSlide(presentation) {
  const item = getSlideContent("q3_screen");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title, "", true);
  addShape(slide, {
    left: 100,
    top: 178,
    width: 1080,
    height: 112,
    fill: COLORS.navy,
    radius: 22,
  });
  addText(slide, "Selected setting", {
    left: 138,
    top: 204,
    width: 260,
    height: 30,
    fontSize: 18,
    color: "#BFD0DB",
  });
  addText(slide, "α1 = 0   |   α2 = 0.35", {
    left: 138,
    top: 242,
    width: 500,
    height: 38,
    fontSize: 29,
    bold: true,
    color: COLORS.white,
  });
  addText(slide, "Best tested feasible boundary", {
    left: 760,
    top: 222,
    width: 360,
    height: 40,
    fontSize: 23,
    bold: true,
    color: "#A9D7D8",
    align: "right",
  });
  const metrics = [
    ["2.123", "urgent mean"],
    ["9.053", "routine mean"],
    ["4.703", "overall mean"],
  ];
  metrics.forEach((metric, i) => {
    addMetric(slide, metric[0], metric[1], {
      left: 148 + i * 342,
      top: 344,
      width: 260,
      color: i === 0 ? COLORS.teal : i === 1 ? COLORS.amber : COLORS.navy,
    });
  });
  addText(slide, item.visibleCopy.warning, {
    left: 220,
    top: 550,
    width: 840,
    height: 52,
    fontSize: 21,
    bold: true,
    color: COLORS.red,
    align: "center",
    valign: "middle",
    fill: COLORS.paleRed,
    radius: 18,
  });
  addFooter(slide, sourceLine(item), 11);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildQ4ScreenSlide(presentation) {
  const item = getSlideContent("q4_screen");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title, "", true);
  addSimpleTable(slide, {
    rows: [
      ["Policy", "Urgent", "Routine", "Overall"],
      ...item.visibleCopy.rows,
    ],
    left: 126,
    top: 164,
    width: 1028,
    height: 304,
    columnWidths: [344, 228, 228, 228],
    riskRows: [3],
  });
  addText(slide, item.visibleCopy.note, {
    left: 126,
    top: 500,
    width: 1028,
    height: 62,
    fontSize: 19,
    bold: true,
    color: COLORS.navy,
    valign: "middle",
    fill: COLORS.paleTeal,
    radius: 16,
    insets: { left: 18, right: 18, top: 8, bottom: 8 },
  });
  addText(slide, item.visibleCopy.warning, {
    left: 126,
    top: 584,
    width: 1028,
    height: 34,
    fontSize: 18,
    bold: true,
    color: COLORS.red,
    align: "center",
  });
  addFooter(slide, sourceLine(item), 12);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildImplementationSlide(presentation) {
  const item = getSlideContent("implementation");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.lightGrey;
  addTitle(slide, item.title, "", true);
  const shared = addShape(slide, {
    left: 198,
    top: 160,
    width: 884,
    height: 92,
    fill: COLORS.navy,
    radius: 22,
    name: "shared-model",
  });
  addText(slide, "SHARED MODEL", {
    left: 232,
    top: 182,
    width: 200,
    height: 26,
    fontSize: 17,
    bold: true,
    color: "#A9D7D8",
  });
  addText(slide, item.visibleCopy.shared.replace("Shared: ", ""), {
    left: 232,
    top: 212,
    width: 816,
    height: 28,
    fontSize: 20,
    color: COLORS.white,
  });
  const labels = [
    ["Q2", "Earliest available slot", COLORS.grey],
    ["Q3", "Backlog-responsive routine postponement", COLORS.teal],
    ["Q4", "Release and OCR rebooking event", COLORS.navy],
  ];
  const boxes = [];
  labels.forEach((label, i) => {
    const left = 82 + i * 394;
    const box = addShape(slide, {
      left,
      top: 360,
      width: 330,
      height: 150,
      fill: COLORS.white,
      lineFill: label[2],
      lineWidth: 2,
      radius: 20,
      name: `implementation-${label[0]}`,
    });
    boxes.push(box);
    addText(slide, label[0], {
      left: left + 24,
      top: 384,
      width: 72,
      height: 44,
      fontSize: 31,
      bold: true,
      color: label[2],
    });
    addText(slide, label[1], {
      left: left + 24,
      top: 438,
      width: 282,
      height: 54,
      fontSize: 19,
      bold: true,
      color: COLORS.ink,
    });
  });
  boxes.forEach((box) => {
    slide.shapes.connect(shared, box, {
      kind: "elbow",
      fromSide: "bottom",
      toSide: "top",
      line: { style: "solid", fill: COLORS.teal, width: 2 },
      tail: { type: "arrow", width: "sm", length: "sm" },
    });
  });
  addText(slide, item.visibleCopy.verification, {
    left: 166,
    top: 560,
    width: 948,
    height: 54,
    fontSize: 20,
    bold: true,
    color: COLORS.navy,
    align: "center",
    valign: "middle",
  });
  addFooter(slide, sourceLine(item), 13);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

function buildLiteratureSlide(presentation) {
  const item = getSlideContent("literature_ai");
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.white;
  addTitle(slide, item.title, "", true);
  addText(slide, "How the literature shaped the model", {
    left: 84,
    top: 170,
    width: 650,
    height: 42,
    fontSize: 26,
    bold: true,
    color: COLORS.navy,
  });
  item.visibleCopy.literature.forEach((line, index) => {
    addShape(slide, {
      geometry: "ellipse",
      left: 88,
      top: 238 + index * 86,
      width: 18,
      height: 18,
      fill: COLORS.teal,
    });
    addText(slide, line, {
      left: 126,
      top: 226 + index * 86,
      width: 610,
      height: 62,
      fontSize: 19,
      color: COLORS.ink,
      valign: "middle",
    });
  });
  addShape(slide, {
    left: 782,
    top: 170,
    width: 398,
    height: 346,
    fill: COLORS.lightGrey,
    radius: 24,
  });
  addText(slide, "AI assistance", {
    left: 820,
    top: 210,
    width: 322,
    height: 42,
    fontSize: 27,
    bold: true,
    color: COLORS.teal,
  });
  addText(slide, item.visibleCopy.ai, {
    left: 820,
    top: 278,
    width: 322,
    height: 178,
    fontSize: 21,
    color: COLORS.ink,
  });
  addText(slide, "Full bibliography: nine peer-reviewed sources in the submitted report.", {
    left: 132,
    top: 564,
    width: 1016,
    height: 50,
    fontSize: 20,
    bold: true,
    color: COLORS.navy,
    align: "center",
    valign: "middle",
    fill: COLORS.paleTeal,
    radius: 20,
  });
  addFooter(slide, sourceLine(item), 14);
  addSpeakerNotes(slide, item.speakerNotes);
  return slide;
}

async function writeBlob(outputPath, blob) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, new Uint8Array(await blob.arrayBuffer()));
}

export async function buildDeck() {
  const presentation = Presentation.create({
    slideSize: { width: SLIDE_W, height: SLIDE_H },
  });
  buildTitleSlide(presentation);
  buildProblemSlide(presentation);
  buildComparisonSlide(presentation);
  buildMethodSlide(presentation);
  buildQ2Slide(presentation);
  buildQ3Slide(presentation);
  buildQ4Slide(presentation);
  buildRecommendationSlide(presentation);
  buildCITableSlide(presentation);
  buildWarmupSlide(presentation);
  buildQ3ScreenSlide(presentation);
  buildQ4ScreenSlide(presentation);
  buildImplementationSlide(presentation);
  buildLiteratureSlide(presentation);
  return presentation;
}

async function main() {
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  await fs.mkdir(path.dirname(OUTPUT), { recursive: true });

  const presentation = await buildDeck();
  if (presentation.slides.items.length !== 14) {
    throw new Error(`Expected 14 slides, got ${presentation.slides.items.length}`);
  }

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(path.join(PREVIEW, `${stem}.png`), png);
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(LAYOUT, `${stem}.json`), await layout.text());
  }

  const montage = await presentation.export({
    format: "webp",
    montage: true,
    scale: 1,
  });
  await writeBlob(path.join(TMP, "deck-montage.webp"), montage);

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT);
  console.log(`DECK BUILD PASS: ${OUTPUT}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
