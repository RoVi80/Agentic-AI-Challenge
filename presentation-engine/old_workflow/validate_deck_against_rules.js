const fs = require("fs");

function loadJson(path) {
  try {
    return JSON.parse(fs.readFileSync(path, "utf8"));
  } catch (err) {
    return {
      __load_error: `Could not read or parse JSON file "${path}": ${err.message}`
    };
  }
}

function countWords(text) {
  if (typeof text !== "string") return 0;
  return text.trim().split(/\s+/).filter(Boolean).length;
}

function countBulletsInContentBlocks(contentBlocks) {
  if (!Array.isArray(contentBlocks)) return 0;

  let count = 0;

  for (const block of contentBlocks) {
    if (Array.isArray(block.bullets)) {
      count += block.bullets.length;
    }

    if (Array.isArray(block.items)) {
      count += block.items.length;
    }

    if (Array.isArray(block.points)) {
      count += block.points.length;
    }
  }

  return count;
}

function addError(errors, message) {
  errors.push(message);
}

function addWarning(warnings, message) {
  warnings.push(message);
}

const presentationPath = process.argv[2];
const rulesPath = process.argv[3];

if (!presentationPath || !rulesPath) {
  console.error("Usage: node src/validate_deck_against_rules.js <presentation.json> <brand_validation_rules.json>");
  process.exit(1);
}

const presentation = loadJson(presentationPath);
const rules = loadJson(rulesPath);

const errors = [];
const warnings = [];

if (presentation.__load_error) addError(errors, presentation.__load_error);
if (rules.__load_error) addError(errors, rules.__load_error);

if (errors.length === 0) {
  if (!Array.isArray(presentation.slides)) {
    addError(errors, "Presentation must contain a slides array.");
  }

  if (!rules.hard_rules || typeof rules.hard_rules !== "object") {
    addError(errors, "Rules must contain hard_rules object.");
  }
}

if (errors.length === 0) {
  const hard = rules.hard_rules;

  for (const slide of presentation.slides) {
    const slideId = slide.id || "unknown";

    if (!hard.allowed_slide_types.includes(slide.type)) {
      addError(
        errors,
        `Slide ${slideId}: slide type "${slide.type}" is not allowed. Allowed types: ${hard.allowed_slide_types.join(", ")}.`
      );
    }

    const titleWords = countWords(slide.title);
    if (titleWords > hard.max_title_words) {
      addError(
        errors,
        `Slide ${slideId}: title has ${titleWords} words, but max_title_words is ${hard.max_title_words}.`
      );
    }

    const bulletCount = countBulletsInContentBlocks(slide.content_blocks);
    if (bulletCount > hard.max_bullets_per_slide) {
      addWarning(
        warnings,
        `Slide ${slideId}: has ${bulletCount} bullet-like items, recommended maximum is ${hard.max_bullets_per_slide}.`
      );
    }

    if (hard.require_footer && (!slide.footer && !presentation.footer)) {
      addWarning(
        warnings,
        `Slide ${slideId}: footer is required by rules, but no slide-level or presentation-level footer was found.`
      );
    }

    if (hard.require_logo && (!slide.logo && !presentation.logo)) {
      addError(
        errors,
        `Slide ${slideId}: logo is required by rules, but no slide-level or presentation-level logo was found.`
      );
    }
  }
}

const result = {
  valid: errors.length === 0,
  status: errors.length === 0 ? "VALID" : "INVALID",
  errors,
  warnings
};

console.log(JSON.stringify(result, null, 2));

process.exit(errors.length === 0 ? 0 : 1);