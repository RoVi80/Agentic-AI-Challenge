const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function loadJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  return JSON.parse(raw);
}

function getTheme(spec) {
  if (spec.theme) return spec.theme;
  if (spec.global_style) return spec.global_style;
  return {};
}

function getColor(theme, key, fallback) {
  return theme?.color_tokens?.[key] || fallback;
}

function getFont(theme, key, fallback) {
  return theme?.font_tokens?.[key] || fallback;
}

function getBlocksByType(slideSpec, blockType) {
  return (slideSpec.content_blocks || []).filter(b => b.block_type === blockType);
}

function getFirstBlockByRole(slideSpec, role) {
  return (slideSpec.content_blocks || []).find(b => b.role === role);
}

function addFooter(slide, slideNumber, theme) {
  const footerText = theme?.footer_text || "";
  if (footerText) {
    slide.addText(footerText, {
      x: 0.3,
      y: 7.0,
      w: 5.0,
      h: 0.2,
      fontSize: 9,
      color: getColor(theme, "muted_text", "666666"),
      margin: 0
    });
  }

  if (theme?.show_slide_numbers) {
    slide.addText(String(slideNumber), {
      x: 12.5,
      y: 7.0,
      w: 0.4,
      h: 0.2,
      fontSize: 9,
      align: "right",
      color: getColor(theme, "muted_text", "666666"),
      margin: 0
    });
  }
}

function renderAccentBar(slide, slideSpec, theme) {
  const bars = (slideSpec.visual_elements || []).filter(
    el => el.element_type === "accent_bar" && el.position === "top"
  );

  if (bars.length > 0) {
    const bar = bars[0];
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.12,
      fill: { color: getColor(theme, bar.color_token || "primary", "2F75B5") },
      line: { color: getColor(theme, bar.color_token || "primary", "2F75B5") }
    });
  }
}

function baseSlide(ppt, theme, slideSpec) {
  const slide = ppt.addSlide();

  const bgType = slideSpec?.background?.type || "solid";
  if (bgType === "solid") {
    slide.background = {
      color: getColor(theme, slideSpec?.background?.color_token || "background", "FFFFFF")
    };
  } else {
    slide.background = {
      color: getColor(theme, "background", "FFFFFF")
    };
  }

  renderAccentBar(slide, slideSpec, theme);
  return slide;
}

function titleStyle(theme, slideSpec) {
  return {
    fontFace: getFont(theme, "title_font", "Aptos"),
    bold: true,
    color: getColor(theme, "text", "1F1F1F"),
    align: slideSpec?.style_overrides?.title_alignment || "left"
  };
}

function bodyStyle(theme) {
  return {
    fontFace: getFont(theme, "body_font", "Aptos"),
    color: getColor(theme, "text", "333333")
  };
}

function renderSlideHeading(slide, slideSpec, theme) {
  const t = titleStyle(theme, slideSpec);
  const titleSize = slideSpec?.style_overrides?.title_size || getFont(theme, "title_size", 24);

  slide.addText(slideSpec.title || "", {
    x: 0.8,
    y: 0.55,
    w: 10.8,
    h: 0.5,
    fontSize: titleSize,
    ...t
  });

  if (slideSpec.subtitle) {
    slide.addText(slideSpec.subtitle, {
      x: 0.8,
      y: 1.05,
      w: 10.0,
      h: 0.25,
      fontSize: 11,
      color: getColor(theme, "muted_text", "666666")
    });
  }
}

function renderTitleSlide(slide, s, theme) {
  const t = titleStyle(theme, s);
  const b = bodyStyle(theme);

  const titleSize = s?.style_overrides?.title_size || 28;
  const subtitleSize = getFont(theme, "subtitle_size", 16);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 1.2,
    w: 11.0,
    h: 0.8,
    fontSize: titleSize,
    ...t
  });

  if (s.subtitle) {
    slide.addText(s.subtitle, {
      x: 0.8,
      y: 2.15,
      w: 10.0,
      h: 0.45,
      fontSize: subtitleSize,
      ...b
    });
  }

  const taglineBlock = getFirstBlockByRole(s, "tagline");
  if (taglineBlock?.text) {
    slide.addText(taglineBlock.text, {
      x: 0.8,
      y: 2.85,
      w: 9.0,
      h: 0.35,
      fontSize: 12,
      italic: true,
      color: getColor(theme, "muted_text", "666666")
    });
  }
}

function renderBulletSummary(slide, s, theme) {
  const b = bodyStyle(theme);
  const bodySize = s?.style_overrides?.body_size || 18;

  renderSlideHeading(slide, s, theme);

  const introBlock = getFirstBlockByRole(s, "intro");
  if (introBlock?.text) {
    slide.addText(introBlock.text, {
      x: 0.9,
      y: 1.55,
      w: 11.0,
      h: 0.4,
      fontSize: 13,
      ...b
    });
  }

  const calloutBlock = getFirstBlockByRole(s, "highlight");
  let bulletY = 2.0;

  if (calloutBlock?.text) {
    const useCard = s?.style_overrides?.surface_style === "card";

    if (useCard) {
      slide.addShape("roundRect", {
        x: 0.9,
        y: 1.8,
        w: 11.2,
        h: 0.8,
        rectRadius: 0.05,
        fill: { color: getColor(theme, "secondary", "EAF2FF") },
        line: { color: getColor(theme, "secondary", "EAF2FF") }
      });
    }

    slide.addText(calloutBlock.text, {
      x: 1.15,
      y: 2.02,
      w: 10.6,
      h: 0.28,
      fontSize: 13,
      bold: true,
      color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5"),
      align: "left"
    });

    bulletY = 2.95;
  }

  const bulletBlock =
    getFirstBlockByRole(s, "main_points") ||
    getFirstBlockByRole(s, "supporting_points") ||
    getBlocksByType(s, "bullet_list")[0];

  const bulletItems = (bulletBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(bulletItems, {
    x: 1.0,
    y: bulletY,
    w: 10.8,
    h: 3.5,
    fontSize: bodySize,
    paraSpaceAfterPt: 10,
    ...b
  });
}

function renderComparison(slide, s, theme) {
  renderSlideHeading(slide, s, theme);

  const leftBlock =
    getFirstBlockByRole(s, "left") ||
    getBlocksByType(s, "comparison_side")[0];

  const rightBlock =
    getFirstBlockByRole(s, "right") ||
    getBlocksByType(s, "comparison_side")[1];

  const bottomCallout = getFirstBlockByRole(s, "bottom_note");

  const leftX = 0.85;
  const rightX = 6.95;
  const cardY = 1.7;
  const cardW = 5.45;
  const cardH = 3.9;

  slide.addShape("roundRect", {
    x: leftX,
    y: cardY,
    w: cardW,
    h: cardH,
    rectRadius: 0.06,
    fill: { color: "F5F5F5" },
    line: { color: "DADADA", pt: 1 }
  });

  slide.addShape("roundRect", {
    x: rightX,
    y: cardY,
    w: cardW,
    h: cardH,
    rectRadius: 0.06,
    fill: { color: getColor(theme, "secondary", "EAF2FF") },
    line: { color: "B7D1FF", pt: 1 }
  });

  slide.addText(leftBlock?.title || "Current state", {
    x: leftX + 0.25,
    y: cardY + 0.22,
    w: cardW - 0.5,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: "333333"
  });

  slide.addText(rightBlock?.title || "Target state", {
    x: rightX + 0.25,
    y: cardY + 0.22,
    w: cardW - 0.5,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: getColor(theme, "text", "1F1F1F")
  });

  const leftItems = (leftBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  const rightItems = (rightBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(leftItems, {
    x: leftX + 0.25,
    y: cardY + 0.72,
    w: cardW - 0.45,
    h: 2.7,
    fontSize: 15,
    color: "444444",
    paraSpaceAfterPt: 8
  });

  slide.addText(rightItems, {
    x: rightX + 0.25,
    y: cardY + 0.72,
    w: cardW - 0.45,
    h: 2.7,
    fontSize: 15,
    color: "222222",
    paraSpaceAfterPt: 8
  });

  if (bottomCallout?.text) {
    slide.addShape("roundRect", {
      x: 1.35,
      y: 5.95,
      w: 10.6,
      h: 0.65,
      rectRadius: 0.05,
      fill: { color: getColor(theme, "surface", "F7FAFE") },
      line: { color: getColor(theme, "secondary", "EAF2FF") }
    });

    slide.addText(bottomCallout.text, {
      x: 1.6,
      y: 6.14,
      w: 10.1,
      h: 0.22,
      fontSize: 11,
      bold: true,
      color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5"),
      align: "center"
    });
  }
}

function renderProcessFlow(slide, s, theme) {
  renderSlideHeading(slide, s, theme);

  const stepBlocks = getBlocksByType(s, "process_step");
  const fallbackSteps = s.content?.steps || [];
  const steps =
    stepBlocks.length > 0
      ? stepBlocks.map(b => ({ label: b.title || b.label, description: b.text || b.description }))
      : fallbackSteps;

  const count = Math.max(steps.length, 1);
  const boxW = 2.15;
  const gap = 0.35;
  const totalW = count * boxW + (count - 1) * gap;
  let startX = Math.max((13.33 - totalW) / 2, 0.5);
  const y = 2.45;

  steps.forEach((step, i) => {
    slide.addShape("roundRect", {
      x: startX,
      y,
      w: boxW,
      h: 1.75,
      rectRadius: 0.07,
      fill: { color: i % 2 === 0 ? getColor(theme, "secondary", "EAF2FF") : "F5F5F5" },
      line: { color: "CFCFCF", pt: 1 }
    });

    slide.addShape("ellipse", {
      x: startX + 0.8,
      y: y - 0.32,
      w: 0.55,
      h: 0.55,
      fill: { color: getColor(theme, "primary", "2F75B5") },
      line: { color: getColor(theme, "primary", "2F75B5") }
    });

    slide.addText(String(i + 1), {
      x: startX + 0.92,
      y: y - 0.15,
      w: 0.3,
      h: 0.15,
      fontSize: 10,
      bold: true,
      align: "center",
      color: "FFFFFF"
    });

    slide.addText(step.label || "", {
      x: startX + 0.12,
      y: y + 0.22,
      w: boxW - 0.24,
      h: 0.25,
      align: "center",
      fontSize: 15,
      bold: true,
      color: getColor(theme, "text", "1F1F1F")
    });

    slide.addText(step.description || "", {
      x: startX + 0.15,
      y: y + 0.62,
      w: boxW - 0.3,
      h: 0.8,
      align: "center",
      valign: "mid",
      fontSize: 10,
      color: "444444"
    });

    if (i < steps.length - 1) {
      slide.addText("→", {
        x: startX + boxW + 0.05,
        y: y + 0.58,
        w: 0.25,
        h: 0.25,
        fontSize: 20,
        color: "666666",
        align: "center"
      });
    }

    startX += boxW + gap;
  });
}

function renderKpiDashboard(slide, s, theme) {
  renderSlideHeading(slide, s, theme);

  const statBlocks = getBlocksByType(s, "stat");
  const fallbackKpis = s.content?.kpis || [];
  const kpis =
    statBlocks.length > 0
      ? statBlocks.map(b => ({
          label: b.label || b.title,
          value: b.value,
          comment: b.text || b.comment
        }))
      : fallbackKpis;

  const count = Math.max(kpis.length, 1);
  const boxW = 3.35;
  const gap = 0.42;
  const totalW = count * boxW + (count - 1) * gap;
  let startX = Math.max((13.33 - totalW) / 2, 0.5);
  const y = 2.1;

  kpis.forEach((kpi, index) => {
    const fillColor = index === 1
      ? getColor(theme, "secondary", "EAF2FF")
      : getColor(theme, "surface", "F7FAFE");

    slide.addShape("roundRect", {
      x: startX,
      y,
      w: boxW,
      h: 2.85,
      rectRadius: 0.08,
      fill: { color: fillColor },
      line: { color: "D9E8FB", pt: 1 }
    });

    slide.addText(kpi.label || "", {
      x: startX + 0.15,
      y: y + 0.22,
      w: boxW - 0.3,
      h: 0.28,
      fontSize: 14,
      bold: true,
      align: "center",
      color: getColor(theme, "text", "1F1F1F")
    });

    slide.addText(kpi.value || "", {
      x: startX + 0.15,
      y: y + 0.82,
      w: boxW - 0.3,
      h: 0.68,
      fontSize: 26,
      bold: true,
      align: "center",
      color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5")
    });

    slide.addText(kpi.comment || "", {
      x: startX + 0.18,
      y: y + 1.8,
      w: boxW - 0.36,
      h: 0.55,
      fontSize: 10,
      align: "center",
      color: "555555"
    });

    startX += boxW + gap;
  });
}

function renderClosing(slide, s, theme) {
  const bodySize = s?.style_overrides?.body_size || 18;

  renderSlideHeading(slide, s, theme);

  const bulletBlock =
    getFirstBlockByRole(s, "summary_points") ||
    getBlocksByType(s, "bullet_list")[0];

  const points = (bulletBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(points, {
    x: 1.0,
    y: 1.9,
    w: 7.0,
    h: 3.5,
    fontSize: bodySize,
    color: getColor(theme, "text", "333333"),
    paraSpaceAfterPt: 8
  });

  const ctaBlock = getFirstBlockByRole(s, "call_to_action");

  if (ctaBlock?.text) {
    slide.addShape("roundRect", {
      x: 8.65,
      y: 2.15,
      w: 3.55,
      h: 1.75,
      rectRadius: 0.08,
      fill: { color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5") },
      line: { color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5") }
    });

    slide.addText(ctaBlock.text, {
      x: 8.9,
      y: 2.76,
      w: 3.05,
      h: 0.35,
      fontSize: 18,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "mid"
    });
  }
}

function renderFallbackSlide(slide, s, theme) {
  slide.addText(s.title || "Unsupported slide", {
    x: 0.8,
    y: 0.8,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    bold: true,
    color: getColor(theme, "text", "1F1F1F")
  });

  slide.addText(`Type not yet implemented: ${s.type}`, {
    x: 0.9,
    y: 1.7,
    w: 10.0,
    h: 0.4,
    fontSize: 14,
    color: getColor(theme, "muted_text", "666666")
  });
}

function renderSlide(ppt, slideSpec, theme, slideNumber) {
  const slide = baseSlide(ppt, theme, slideSpec);

  switch (slideSpec.type) {
    case "title":
      renderTitleSlide(slide, slideSpec, theme);
      break;
    case "bullet_summary":
      renderBulletSummary(slide, slideSpec, theme);
      break;
    case "comparison":
      renderComparison(slide, slideSpec, theme);
      break;
    case "process_flow":
      renderProcessFlow(slide, slideSpec, theme);
      break;
    case "kpi_dashboard":
      renderKpiDashboard(slide, slideSpec, theme);
      break;
    case "closing":
      renderClosing(slide, slideSpec, theme);
      break;
    default:
      renderFallbackSlide(slide, slideSpec, theme);
  }

  addFooter(slide, slideNumber, theme);
}

function mergeThemeWithProfile(theme, profile) {
  if (!profile) return theme || {};

  return {
    ...theme,

    // override colors
    color_tokens: {
      ...(theme?.color_tokens || {}),
      ...(profile?.colors || {})
    },

    // override fonts
    font_tokens: {
      ...(theme?.font_tokens || {}),
      title_font: profile?.typography?.title_font || theme?.font_tokens?.title_font,
      body_font: profile?.typography?.body_font || theme?.font_tokens?.body_font,
      caption_font: profile?.typography?.caption_font || theme?.font_tokens?.caption_font
    },

    // footer rules
    footer_text: profile?.footer?.footer_text || theme?.footer_text,
    show_slide_numbers:
      profile?.footer?.slide_numbers_required ?? theme?.show_slide_numbers,

    // branding
    brand_name: profile?.profile_name || theme?.brand_name,

    // spacing
    spacing_tokens: {
      ...(theme?.spacing_tokens || {}),
      ...(profile?.spacing || {})
    }
  };
}

// helper functions for validation

function countWords(text) {
  if (!text || typeof text !== "string") return 0;
  return text.trim().split(/\s+/).filter(Boolean).length;
}

function getBulletItems(slideSpec) {
  const blocks = slideSpec.content_blocks || [];
  const bulletBlock = blocks.find(b => b.block_type === "bullet_list");
  return bulletBlock?.items || [];
}

// validation function
function validateSlide(slideSpec, profile) {
  const warnings = [];

  const allowedSlideTypes = profile?.allowed_slide_types || [];
  if (
    allowedSlideTypes.length > 0 &&
    slideSpec.type &&
    !allowedSlideTypes.includes(slideSpec.type)
  ) {
    warnings.push(
      `Unsupported slide type "${slideSpec.type}". Allowed types: ${allowedSlideTypes.join(", ")}`
    );
  }

  const titleRequired = profile?.title_rules?.title_required_on_all_slides;
  if (titleRequired && (!slideSpec.title || !slideSpec.title.trim())) {
    warnings.push("Missing required slide title.");
  }

  const maxTitleWords = profile?.title_rules?.max_title_words;
  if (maxTitleWords && countWords(slideSpec.title) > maxTitleWords) {
    warnings.push(
      `Title too long: ${countWords(slideSpec.title)} words (max ${maxTitleWords}).`
    );
  }

  const maxBullets = profile?.density_rules?.max_bullets_per_slide;
  const bulletItems = getBulletItems(slideSpec);
  if (maxBullets && bulletItems.length > maxBullets) {
    warnings.push(
      `Too many bullets: ${bulletItems.length} bullets (max ${maxBullets}).`
    );
  }

  return warnings;
}

function validateDeck(spec, profile) {
  const allWarnings = [];

  (spec.slides || []).forEach((slideSpec, index) => {
    const slideWarnings = validateSlide(slideSpec, profile);

    slideWarnings.forEach(warning => {
      allWarnings.push({
        slideId: slideSpec.id || `slide_${index + 1}`,
        slideNumber: index + 1,
        warning
      });
    });
  });

  return allWarnings;
}

async function main() {
  const inputJson = process.argv[2];
  const outputPptx = process.argv[3] || "output.pptx";
  const profilePath = process.argv[4]; // NEW

  assert(inputJson, "Usage: node render_presentation.js <input.json> <output.pptx> [profile.json]");

  const spec = loadJson(inputJson);

  let profile = {};
  if (profilePath) {
    console.log(`Loading profile: ${profilePath}`);
    profile = loadJson(profilePath);
  }

  assert(
    spec.schema_version === "smart_presentation_v1" ||
      spec.schema_version === "smart_presentation_v2",
    "schema_version must be smart_presentation_v1 or smart_presentation_v2"
  );

  assert(Array.isArray(spec.slides), "slides must be an array");

  const theme = mergeThemeWithProfile(getTheme(spec), profile);
  const validationWarnings = validateDeck(spec, profile);

  if (validationWarnings.length > 0) {
    console.log("\nValidation warnings:");
    validationWarnings.forEach(item => {
      console.log(
        `- Slide ${item.slideNumber} (${item.slideId}): ${item.warning}`
      );
    });
    console.log("");
  } else {
    console.log("Validation passed: no warnings found.");
  }

  const ppt = new PptxGenJS();
  ppt.layout = "LAYOUT_WIDE";
  ppt.author = "Smart PPT Renderer";
  ppt.company = theme?.brand_name || profile?.profile_name || "Unknown";
  ppt.subject = spec.deck?.purpose || "";
  ppt.title = spec.deck?.title || "Presentation";
  ppt.lang = spec.deck?.language || "en-US";

  spec.slides.forEach((slideSpec, index) => {
    renderSlide(ppt, slideSpec, theme, index + 1);
  });

  await ppt.writeFile({ fileName: outputPptx });
  console.log(`PowerPoint created: ${outputPptx}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});