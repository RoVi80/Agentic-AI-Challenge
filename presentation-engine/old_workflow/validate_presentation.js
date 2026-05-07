const fs = require("fs");

function fail(msg) {
  console.error(`INVALID: ${msg}`);
  process.exit(1);
}

function loadJson(path) {
  return JSON.parse(fs.readFileSync(path, "utf8"));
}

const filePath = process.argv[2];
if (!filePath) {
  fail("Usage: node src/validate_presentation.js <presentation.json>");
}

const data = loadJson(filePath);

if (data.schema_version !== "smart_presentation_v2") {
  fail("schema_version must be smart_presentation_v2");
}

if (!Array.isArray(data.slides) || data.slides.length === 0) {
  fail("slides must be a non-empty array");
}

const requiredSlideKeys = [
  "id",
  "section_id",
  "type",
  "variant",
  "title",
  "subtitle",
  "objective",
  "key_message",
  "importance",
  "audience_takeaway",
  "content_blocks",
  "visual_elements",
  "layout",
  "background",
  "style_overrides",
  "notes",
  "source_refs",
  "flags"
];

for (const slide of data.slides) {
  for (const key of requiredSlideKeys) {
    if (!(key in slide)) {
      fail(`Slide ${slide.id || "unknown"} missing key: ${key}`);
    }
  }
}

console.log("VALID: presentation passed basic validation.");