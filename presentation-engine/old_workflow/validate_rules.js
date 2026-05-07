const fs = require("fs");

function fail(msg) {
  console.error(`INVALID: ${msg}`);
  process.exit(1);
}

function loadJson(path) {
  try {
    return JSON.parse(fs.readFileSync(path, "utf8"));
  } catch (err) {
    fail(`Could not read or parse JSON: ${err.message}`);
  }
}

function isStringArray(value) {
  return Array.isArray(value) && value.every(v => typeof v === "string");
}

function isBoolean(value) {
  return typeof value === "boolean";
}

function isNumber(value) {
  return typeof value === "number" && !Number.isNaN(value);
}

const filePath = process.argv[2];

if (!filePath) {
  fail("Usage: node src/validate_rules.js <brand_validation_rules.json>");
}

const rules = loadJson(filePath);

if (rules.rules_version !== "brand_validation_rules_v1") {
  fail("rules_version must be brand_validation_rules_v1");
}

const requiredTopLevel = [
  "rules_version",
  "brand_name",
  "source_profile_name",
  "hard_rules",
  "soft_rules",
  "validation_behavior"
];

for (const key of requiredTopLevel) {
  if (!(key in rules)) {
    fail(`Missing top-level key: ${key}`);
  }
}

const hard = rules.hard_rules;

if (!isStringArray(hard.allowed_colors)) {
  fail("hard_rules.allowed_colors must be an array of strings");
}

if (!isStringArray(hard.allowed_fonts)) {
  fail("hard_rules.allowed_fonts must be an array of strings");
}

if (!isStringArray(hard.allowed_slide_types) || hard.allowed_slide_types.length === 0) {
  fail("hard_rules.allowed_slide_types must be a non-empty array of strings");
}

if (!isNumber(hard.max_bullets_per_slide)) {
  fail("hard_rules.max_bullets_per_slide must be a number");
}

if (!isNumber(hard.max_title_words)) {
  fail("hard_rules.max_title_words must be a number");
}

if (!isBoolean(hard.require_footer)) {
  fail("hard_rules.require_footer must be boolean");
}

if (!isBoolean(hard.require_logo)) {
  fail("hard_rules.require_logo must be boolean");
}

if (!hard.logo_rules || typeof hard.logo_rules !== "object") {
  fail("hard_rules.logo_rules must be an object");
}

const logo = hard.logo_rules;

for (const key of [
  "allow_recolor",
  "allow_rotation",
  "allow_distortion",
  "allow_effects",
  "respect_protection_area"
]) {
  if (!isBoolean(logo[key])) {
    fail(`hard_rules.logo_rules.${key} must be boolean`);
  }
}

if (!isStringArray(logo.allowed_backgrounds)) {
  fail("hard_rules.logo_rules.allowed_backgrounds must be an array of strings");
}

if (!rules.soft_rules || typeof rules.soft_rules !== "object") {
  fail("soft_rules must be an object");
}

if (!rules.soft_rules.preferred_colors || typeof rules.soft_rules.preferred_colors !== "object") {
  fail("soft_rules.preferred_colors must be an object");
}

if (!isStringArray(rules.soft_rules.warnings_only)) {
  fail("soft_rules.warnings_only must be an array of strings");
}

const behavior = rules.validation_behavior;

if (behavior.hard_rule_violation !== "fail") {
  fail("validation_behavior.hard_rule_violation must be fail");
}

if (behavior.soft_rule_violation !== "warn") {
  fail("validation_behavior.soft_rule_violation must be warn");
}

if (behavior.missing_optional_information !== "warn") {
  fail("validation_behavior.missing_optional_information must be warn");
}

if (behavior.missing_required_information !== "fail") {
  fail("validation_behavior.missing_required_information must be fail");
}

console.log("VALID: validation rules passed basic validation.");