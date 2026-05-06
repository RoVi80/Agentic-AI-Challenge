// tools/simple_powerpoint_renderer.js
// A clean-start local PowerPoint renderer for the Agentic AI Challenge.
//
// Purpose:
// - Takes one simple JSON file as input.
// - Renders a .pptx file with PptxGenJS.
// - Supports basic user-controllable settings:
//   content, slide amount, font, font color, background color, accent color.
//
// Run locally:
//   node tools/simple_powerpoint_renderer.js input/simple_deck.json output/simple_demo.pptx
//
// JSON shape expected:
// {
//   "title": "Deck title",
//   "subtitle": "Optional subtitle",
//   "settings": {
//     "fontFace": "Arial",
//     "titleColor": "24135F",
//     "bodyColor": "1D1D1D",
//     "backgroundColor": "FFFFFF",
//     "accentColor": "D0006F",
//     "footerText": "Optional footer"
//   },
//   "slides": [
//     {
//       "type": "title",
//       "title": "Main title",
//       "subtitle": "Short subtitle"
//     },
//     {
//       "type": "bullets",
//       "title": "Slide title",
//       "bullets": ["Point one", "Point two", "Point three"]
//     },
//     {
//       "type": "two_column",
//       "title": "Comparison",
//       "leftTitle": "Before",
//       "leftBullets": ["Manual", "Slow"],
//       "rightTitle": "After",
//       "rightBullets": ["Automated", "Faster"]
//     },
//     {
//       "type": "closing",
//       "title": "Conclusion",
//       "bullets": ["Key takeaway", "Next step"]
//     }
//   ]
// }

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const DEFAULTS = {
  fontFace: "Arial",
  titleColor: "24135F",
  bodyColor: "1D1D1D",
  backgroundColor: "FFFFFF",
  accentColor: "D0006F",
  mutedColor: "B1B3B3",
  footerText: "Generated presentation"
};

function usageAndExit() {
  console.error("Usage: node tools/simple_powerpoint_renderer.js <input.json> <output.pptx>");
  process.exit(1);
}

function cleanHex(value, fallback) {
  if (!value || typeof value !== "string") return fallback;
  const cleaned = value.trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(cleaned)) return cleaned;
  return fallback;
}

function safeString(value, fallback = "") {
  if (value === null || value === undefined) return fallback;
  return String(value);
}

function readJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(`Could not parse JSON file '${filePath}': ${err.message}`);
  }
}

function normalizeDeck(deck) {
  if (!deck || typeof deck !== "object") {
    throw new Error("Input JSON must be an object.");
  }

  if (!Array.isArray(deck.slides) || deck.slides.length === 0) {
    throw new Error("Input JSON must contain a non-empty 'slides' array.");
  }

  const rawSettings = deck.settings || {};

  const settings = {
    fontFace: safeString(rawSettings.fontFace, DEFAULTS.fontFace),
    titleColor: cleanHex(rawSettings.titleColor, DEFAULTS.titleColor),
    bodyColor: cleanHex(rawSettings.bodyColor, DEFAULTS.bodyColor),
    backgroundColor: cleanHex(rawSettings.backgroundColor, DEFAULTS.backgroundColor),
    accentColor: cleanHex(rawSettings.accentColor, DEFAULTS.accentColor),
    mutedColor: cleanHex(rawSettings.mutedColor, DEFAULTS.mutedColor),
    footerText: safeString(rawSettings.footerText, DEFAULTS.footerText)
  };

  return {
    title: safeString(deck.title, "Untitled Presentation"),
    subtitle: safeString(deck.subtitle, ""),
    settings,
    slides: deck.slides
  };
}

function addFooter(slide, slideNumber, settings) {
  slide.addText(settings.footerText, {
    x: 0.55,
    y: 7.08,
    w: 8.5,
    h: 0.2,
    fontFace: settings.fontFace,
    fontSize: 7,
    color: settings.mutedColor,
    margin: 0
  });

  slide.addText(String(slideNumber), {
    x: 12.35,
    y: 7.08,
    w: 0.45,
    h: 0.2,
    fontFace: settings.fontFace,
    fontSize: 7,
    color: settings.mutedColor,
    align: "right",
    margin: 0
  });
}

function addTopBar(slide, pptx, settings) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.16,
    fill: { color: settings.accentColor },
    line: { color: settings.accentColor }
  });
}

function addSlideTitle(slide, pptx, title, settings) {
  addTopBar(slide, pptx, settings);
  slide.addText(safeString(title, "Slide"), {
    x: 0.65,
    y: 0.45,
    w: 11.9,
    h: 0.45,
    fontFace: settings.fontFace,
    fontSize: 23,
    bold: true,
    color: settings.titleColor,
    margin: 0,
    fit: "shrink"
  });
}

function addBulletList(slide, bullets, settings, box) {
  const safeBullets = Array.isArray(bullets) ? bullets.slice(0, 7) : [];

  if (safeBullets.length === 0) {
    slide.addText("No content provided.", {
      ...box,
      fontFace: settings.fontFace,
      fontSize: 15,
      color: settings.bodyColor,
      margin: 0
    });
    return;
  }

  const runs = safeBullets.map((item) => ({
    text: safeString(item),
    options: {
      bullet: { type: "ul" },
      breakLine: true
    }
  }));

  slide.addText(runs, {
    ...box,
    fontFace: settings.fontFace,
    fontSize: 16,
    color: settings.bodyColor,
    margin: 0.05,
    fit: "shrink",
    paraSpaceAfterPt: 10,
    breakLine: false
  });
}

function renderTitleSlide(pptx, spec, settings, slideNumber) {
  const slide = pptx.addSlide();
  slide.background = { color: settings.backgroundColor };

  addTopBar(slide, pptx, settings);

  slide.addText(safeString(spec.title, "Untitled Presentation"), {
    x: 0.8,
    y: 2.35,
    w: 11.6,
    h: 0.8,
    fontFace: settings.fontFace,
    fontSize: 34,
    bold: true,
    color: settings.titleColor,
    margin: 0,
    fit: "shrink"
  });

  const subtitle = safeString(spec.subtitle, "");
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.82,
      y: 3.28,
      w: 10.9,
      h: 0.45,
      fontFace: settings.fontFace,
      fontSize: 16,
      color: settings.bodyColor,
      margin: 0,
      fit: "shrink"
    });
  }

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.8,
    y: 4.2,
    w: 2.2,
    h: 0.08,
    fill: { color: settings.accentColor },
    line: { color: settings.accentColor }
  });

  addFooter(slide, slideNumber, settings);
}

function renderBulletsSlide(pptx, spec, settings, slideNumber) {
  const slide = pptx.addSlide();
  slide.background = { color: settings.backgroundColor };

  addSlideTitle(slide, pptx, spec.title, settings);
  addBulletList(slide, spec.bullets, settings, {
    x: 1.0,
    y: 1.55,
    w: 11.2,
    h: 4.85
  });

  addFooter(slide, slideNumber, settings);
}

function renderTwoColumnSlide(pptx, spec, settings, slideNumber) {
  const slide = pptx.addSlide();
  slide.background = { color: settings.backgroundColor };

  addSlideTitle(slide, pptx, spec.title, settings);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 1.35,
    w: 5.75,
    h: 5.2,
    rectRadius: 0.08,
    fill: { color: "F7F7F7" },
    line: { color: settings.mutedColor, width: 0.7 }
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.8,
    y: 1.35,
    w: 5.75,
    h: 5.2,
    rectRadius: 0.08,
    fill: { color: "F7F7F7" },
    line: { color: settings.mutedColor, width: 0.7 }
  });

  slide.addText(safeString(spec.leftTitle, "Option A"), {
    x: 1.1,
    y: 1.65,
    w: 5.1,
    h: 0.3,
    fontFace: settings.fontFace,
    fontSize: 16,
    bold: true,
    color: settings.titleColor,
    margin: 0
  });

  slide.addText(safeString(spec.rightTitle, "Option B"), {
    x: 7.1,
    y: 1.65,
    w: 5.1,
    h: 0.3,
    fontFace: settings.fontFace,
    fontSize: 16,
    bold: true,
    color: settings.titleColor,
    margin: 0
  });

  addBulletList(slide, spec.leftBullets, settings, {
    x: 1.15,
    y: 2.2,
    w: 4.95,
    h: 3.75
  });

  addBulletList(slide, spec.rightBullets, settings, {
    x: 7.15,
    y: 2.2,
    w: 4.95,
    h: 3.75
  });

  addFooter(slide, slideNumber, settings);
}

function renderClosingSlide(pptx, spec, settings, slideNumber) {
  const slide = pptx.addSlide();
  slide.background = { color: settings.backgroundColor };

  addTopBar(slide, pptx, settings);

  slide.addText(safeString(spec.title, "Thank you"), {
    x: 0.8,
    y: 1.55,
    w: 11.6,
    h: 0.7,
    fontFace: settings.fontFace,
    fontSize: 31,
    bold: true,
    color: settings.titleColor,
    margin: 0,
    fit: "shrink"
  });

  addBulletList(slide, spec.bullets, settings, {
    x: 1.0,
    y: 2.65,
    w: 10.8,
    h: 2.7
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 6.82,
    w: 13.333,
    h: 0.26,
    fill: { color: settings.accentColor },
    line: { color: settings.accentColor }
  });

  addFooter(slide, slideNumber, settings);
}

async function renderPresentation(inputPath, outputPath) {
  const rawDeck = readJson(inputPath);
  const deck = normalizeDeck(rawDeck);

  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Agentic AI Challenge";
  pptx.company = "Agentic AI Challenge";
  pptx.subject = deck.title;
  pptx.title = deck.title;
  pptx.lang = "en-US";

  deck.slides.forEach((slideSpec, index) => {
    const type = safeString(slideSpec.type, "bullets").toLowerCase();
    const slideNumber = index + 1;

    if (type === "title") {
      renderTitleSlide(pptx, slideSpec, deck.settings, slideNumber);
    } else if (type === "two_column" || type === "two-column") {
      renderTwoColumnSlide(pptx, slideSpec, deck.settings, slideNumber);
    } else if (type === "closing") {
      renderClosingSlide(pptx, slideSpec, deck.settings, slideNumber);
    } else {
      renderBulletsSlide(pptx, slideSpec, deck.settings, slideNumber);
    }
  });

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await pptx.writeFile({ fileName: outputPath });

  return {
    success: true,
    outputPath,
    slideCount: deck.slides.length
  };
}

async function main() {
  const [, , inputPath, outputPath] = process.argv;

  if (!inputPath || !outputPath) {
    usageAndExit();
  }

  try {
    const result = await renderPresentation(inputPath, outputPath);
    console.log(JSON.stringify(result, null, 2));
  } catch (err) {
    console.error(JSON.stringify({ success: false, error: err.message }, null, 2));
    process.exit(1);
  }
}

main();
