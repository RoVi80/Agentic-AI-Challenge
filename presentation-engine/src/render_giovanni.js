const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

// ── Layout constants (LAYOUT_WIDE = 13.33" × 7.5") ───────────────────────────
const W  = 13.33;
const M  = 0.5;
const CW = (W - 2 * M - 0.5) / 2;   // column width  ≈ 5.915"
const CR = M + CW + 0.5;             // right-column x ≈ 6.915"
const TW = W - 2 * M;                // full text width ≈ 12.33"

// ── Brand loaders ─────────────────────────────────────────────────────────────

function loadColors(brand) {
  const strip = (v) => (v || "").replace("#", "");
  const p = brand.colors?.primary   ?? [];
  const s = brand.colors?.secondary ?? [];
  
  const out = {}

	if (p.length < 3) {
		out.primary3 = "FFFFFF";
	}
	if (s.length < 2) {
		out.secondary2 = "FFFFFF";
	}

  for (let i = 0; i < p.length; i++) {
    out[`primary${i + 1}`] = strip(p[i]) || `24135F`;
  }
  for (let i = 0; i < s.length; i++) {
    out[`secondary${i + 1}`] = strip(s[i]) || `2DCCD3`;
  }
  out.white = "FFFFFF";
	out.black = "000000";
  out.lightGrey = "b2b2b2";
  return out;
}

function loadFont(brand) {
  return {
    primary:  {
        name: brand.font?.primary  || "Arial",
        weights: brand.font?.weights?.primary || [400, 700],
    },
    fallback: { 
        name: brand.font?.fallback || "Arial",
        weights: brand.font?.weights?.fallback || [400, 700],
    },
  };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function rect(color, x, y, w, h) {
  return { rect: { x, y, w, h, fill: { color }, line: { color, width: 0 } } };
}

function img(filePath, x, y, w, h) {
  return { image: { path: filePath, x, y, w, h } };
}

function addRect(slide, color, x, y, w, h) {
  slide.addShape("rect", { x, y, w, h, fill: { color }, line: { color, width: 0 } });
}

function resolveKeyPoint(kp) {
  if (typeof kp === "string") return { title: kp, slides: [{ left: kp, right: null }] };
  const slides = kp.slides ?? [{ left: kp.left ?? kp.content ?? kp.title ?? "", right: kp.right ?? null }];
  return { title: kp.title ?? "", slides };
}

function distributeContentSlides(sectionCount, desiredTotal) {
  // Fixed overhead: title(1) + agenda(1) + closing(1) + one section slide per section
  const overhead = 3 + sectionCount;
  const contentSlides = desiredTotal != null
    ? Math.max(0, desiredTotal - overhead)
    : sectionCount;
  const base   = Math.floor(contentSlides / sectionCount);
  const extras = contentSlides % sectionCount;
  return Array.from({ length: sectionCount }, (_, i) => base + (i < extras ? 1 : 0));
}

// ── Slide masters ─────────────────────────────────────────────────────────────

function defineMasters(pptx, colors, logos, font) {
	const half = W / 2;

	pptx.defineSlideMaster({
		title: "MASTER_TITLE",
		background: { color: colors.primary1 },
		objects: [
			rect(colors.primary2, 0, 0, 0.18, 7.5),
			...(logos.primary ? [img(logos.primary, 0.4, 0.3, 2.1, 1)] : []),
			{ placeholder: {
				options: {
					name: "title",
					type: "title",
					x: M, y: 2.6, w: TW, h: 1.4,
					fontSize: 40, bold: true, color: colors.white, fontFace: font.primary.name, align: "left",
					isTitle: true,
				}}
			},
			{ placeholder: {
					options: {
							name: "subtitle",
							type: "body",
							x: M, y: 4.1, w: TW - 1, h: 0.8,
							fontSize: 18, color: colors.lightGrey, fontFace: font.primary.name, align: "left",
					}}
    },
    { placeholder: {
        options: {
            name: "date",
            type: "date",
            x: M, y: 6.9, w: TW, h: 0.4,
            fontSize: 10, color: colors.lightGrey, fontFace: font.primary.name, align: "left",
        }}
    }
    ],
  });

  const ContentMaster = {
    title: "MASTER_CONTENT",
    objects: [
      rect(colors.primary1, 0, 0, W, 0.12),
      ...(logos.icon ? [img(logos.icon, W - 1.3, 7.5 - 1.3, 1, 1)] : []),
			{ placeholder: {
				options: {
					name: "title",
					type: "title",
					x: M, y: 0.7, w: TW, h: 0.8,
					fontSize: 28, bold: true, color: colors.primary1, fontFace: font.primary.name, align: "left",
					isTitle: true
				}}
			},
			rect(colors.primary2, M, 1.55, TW, 0.025),
			{ placeholder: {
				options: {
					name: "content",
					type: "body",
					x: M, y: 1.6, w: TW, h: 5.5,
					fontSize: 13, color: colors.black, fontFace: font.primary.name, valign: "top", wrap: true
				}}
			},
			{ placeholder: { 
        options: { 
          name: "slideNumber", 
					type: "slideNumber", 
					x: 0.5, y: 7.5, w: 1, h: 0.5,
					align: "left",
					fontSize: 10, fontFace: font.primary.name, color: colors.lightGrey
        } 
    }}
    ],
  };

  pptx.defineSlideMaster(ContentMaster);

  pptx.defineSlideMaster({
    title: "MASTER_SECTION",
    objects: [
      rect(colors.primary1, 0,           0, half,        7.5),
      rect(colors.secondary1,  half,        0, 0.08,        7.5),
      rect(colors.white,   half + 0.08, 0, half - 0.08, 7.5),
      ...(logos.primary ? [img(logos.primary, 0.35, 0.28, 2.1, 1)] : []),
			{ placeholder: {
				options: {
					name: "sectionNumber",
					type: "body",
					x: 0.3, y: 2.2, w: half - 0.5, h: 2.5,
					fontSize: 120, bold: true, color: colors.primary3, fontFace: font.primary.name, transparency: 60,
				}}
			},
			{ placeholder: {
				options: {
					name: "sectionTitle",
					type: "title",
					x: half + 0.3, y: 2.8, w: half - 0.8, h: 1.5,
					fontSize: 26, bold: true, color: colors.primary1, fontFace: font.primary.name, align: "left",
					isTitle: true,
				}}
			},
    ],
  });

  pptx.defineSlideMaster({
    title: "MASTER_CLOSING",
    background: { color: colors.primary1 },
    objects: [
      rect(colors.secondary1,    0, 7.1,  W,    0.4),
      rect(colors.primary2, 0, 0,    0.18, 7.5),
      ...(logos.primary ? [img(logos.primary, 0.4, 0.3, 2.1, 1)] : []),
			{ placeholder: {
				options: {
					name: "title",
					type: "title",
					x: M, y: 2.5, w: TW, h: 1.6,
					fontSize: 38, bold: true, color: colors.white, fontFace: font.primary.name, align: "center",
					isTitle: true,
				}}
			},
			{ placeholder: {
				options: {
					name: "company",
					type: "body",
					x: M, y: 4.4, w: TW, h: 0.6,
					fontSize: 14, color: colors.lightGrey, fontFace: font.primary.name, align: "center",
				}}
			}
    ],
  });

	pptx.defineSlideMaster({
		...ContentMaster,
		title: "MASTER_AGENDA",
		objects: [
			...ContentMaster.objects,
			{ placeholder: {
				options: {
					name: "agendaTitle",
					type: "body",
					x: M, y: 0.25, w: 3, h: 0.4,
					fontSize: 9, bold: true, color: colors.primary2, fontFace: font.primary.name, charSpacing: 3,
				}}
			},
		]
	});
}

// ── Slide builders ────────────────────────────────────────────────────────────

function addTitleSlide(pptx, brief, brand) {
  const slide = pptx.addSlide({ masterName: "MASTER_TITLE" });

  slide.addText(brief.topic ?? "", {placeholder: "title"});

  if (brief.purpose) {
    slide.addText(brief.purpose, {placeholder: "subtitle"});
  }
  const month = new Date().toLocaleDateString("en-UK", { month: "long", year: "numeric" });
  slide.addText(`${month} · ${brand.company_name ?? ""}`, {placeholder: "date"});
}

function addAgendaSlide(pptx, items, colors, font) {
  const slide = pptx.addSlide({ masterName: "MASTER_AGENDA" });

  slide.addText("AGENDA", {placeholder: "agendaTitle"});
	slide.addText("What we'll cover today", {placeholder: "title"});

  items.forEach((kp, i) => {
    const label = resolveKeyPoint(kp).title;
    const y = 1.8 + i * 0.75;
    addRect(slide, colors.primary2, M, y + 0.05, 0.38, 0.38);
    slide.addText(String(i + 1), {
      x: M, y: y + 0.05, w: 0.38, h: 0.38,
      fontSize: 12, bold: true, color: colors.white,
      fontFace: font.primary.name, align: "center", valign: "middle",
    });
    slide.addText(label, {
      x: 1.05, y, w: W - 1.05 - M, h: 0.55,
      fontSize: 16, color: colors.black, fontFace: font.primary.name, valign: "middle",
    });
  });
}

function addSectionSlide(pptx, sectionNumber, sectionTitle) {
  const slide = pptx.addSlide({ masterName: "MASTER_SECTION" });

  slide.addText(String(sectionNumber).padStart(2, "0"), {placeholder: "sectionNumber"});
  slide.addText(sectionTitle, {placeholder: "sectionTitle"});
}

function addContentSlide(pptx, title, text) {
  const slide = pptx.addSlide({ masterName: "MASTER_CONTENT" });

  slide.addText(title, {placeholder: "title"});
	slide.addText(text, {placeholder: "content"});

}

function addClosingSlide(pptx, brand) {
  const slide = pptx.addSlide({ masterName: "MASTER_CLOSING" });

  slide.addText("Thank you", {placeholder: "title"});
  slide.addText(brand.company_name ?? "", {placeholder: "company"});
}

// ── Main generator ────────────────────────────────────────────────────────────

async function generatePresentation(json, outputPath) {
  const brief = json.presentation_brief ?? {};
  const brand = json.brand_guidelines  ?? {};

  const colors = loadColors(brand);
  const font   = loadFont(brand);
  const logos  = {
    primary: brand.logos?.primary || null,
    icon:  brand.logos?.icon  || null,
  };

  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  defineMasters(pptx, colors, logos, font);

  const keyPoints    = brief.key_points ?? [];
  const sections     = keyPoints.map(resolveKeyPoint);
  const distribution = distributeContentSlides(sections.length, brief.desired_slide_count);

  addTitleSlide(pptx, brief, brand);

  if (sections.length > 0) {
    addAgendaSlide(pptx, keyPoints, colors, font);
  }

  sections.forEach((section, i) => {
    addSectionSlide(pptx, i + 1, section.title);
    const count = distribution[i];
    for (let j = 0; j < count; j++) {
      const slide = section.slides[j] ?? { left: "", right: null };
      addContentSlide(pptx, section.title, slide.left);
    }
  });

  addClosingSlide(pptx, brand);

  const out = outputPath ?? path.resolve(__dirname, "presentation.pptx");
  await pptx.writeFile({ fileName: out });
  console.error(`Presentation saved → ${out}`);
  return out;
}

// ── CLI entry point ───────────────────────────────────────────────────────────

if (process.argv[1] === __filename) {
  const [,, inputPath, outputPath] = process.argv;

  if (!inputPath || !outputPath) {
    console.error("Usage: node render_giovanni.js <input.json> <output.pptx>");
    process.exit(1);
  }

  const json = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  generatePresentation(json, outputPath);
}