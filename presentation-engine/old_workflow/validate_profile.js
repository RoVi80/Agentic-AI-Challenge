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
  fail("Usage: node src/validate_profile.js <profile.json>");
}

const profile = loadJson(filePath);

if (profile.profile_version !== "company_style_profile_v1") {
  fail("profile_version must be company_style_profile_v1");
}

const requiredTopLevel = [
  "profile_version",
  "profile_name",
  "description",
  "branding",
  "colors",
  "typography",
  "footer",
  "density_rules",
  "title_rules",
  "allowed_slide_types",
  "validation_rules"
];

for (const key of requiredTopLevel) {
  if (!(key in profile)) {
    fail(`Missing top-level key: ${key}`);
  }
}

if (!Array.isArray(profile.allowed_slide_types) || profile.allowed_slide_types.length === 0) {
  fail("allowed_slide_types must be a non-empty array");
}

if (typeof profile.colors.primary !== "string" || !profile.colors.primary.startsWith("#")) {
  fail("colors.primary must be a hex string");
}

if (typeof profile.typography.title_font !== "string" || !profile.typography.title_font.trim()) {
  fail("typography.title_font must be a non-empty string");
}

if (typeof profile.footer.footer_required !== "boolean") {
  fail("footer.footer_required must be boolean");
}

if (typeof profile.density_rules.max_bullets_per_slide !== "number") {
  fail("density_rules.max_bullets_per_slide must be a number");
}

if (typeof profile.title_rules.max_title_words !== "number") {
  fail("title_rules.max_title_words must be a number");
}

console.log("VALID: profile passed basic validation.");