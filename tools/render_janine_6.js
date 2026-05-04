const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// -------------------------
// Safe helpers (in case JSON values are null)
// TODO: check default values if they make sense
// -------------------------

function safe(value, fallback) {
    return value ?? fallback;
}

function safeNum(value, fallback) {
    return typeof value === "number" && !isNaN(value) ? value : fallback;
}

// -------------------------
// Read JSON from stdin
// -------------------------

function readStdin() {
    return new Promise((resolve, reject) => {
        let data = "";

        process.stdin.setEncoding("utf8");

        process.stdin.on("data", chunk => {
            data += chunk;
        });

        process.stdin.on("end", () => {
            try {
                resolve(JSON.parse(data));
            } catch (err) {
                reject(new Error(`Invalid JSON input: ${err.message}`));
            }
        });

        process.stdin.on("error", reject);
    });
}

// -------------------------
// Main builder
// -------------------------

async function buildPresentation(config) {
    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";

    const outputFile = safe(config.outputFile, path.join(process.cwd(), "output.pptx"));

    // -------------------------
    // Deep-safe config normalization
    // -------------------------

    const colors = safe(config.colors, {});
    const fonts = safe(config.fonts, {});
    const styles = safe(config.styles, {});
    const layout = safe(config.layout, {});
    const bg = safe(config.background, {});

    const safeColors = {
        primary: safe(colors.primary, "FFFFFF"),
        secondary: safe(colors.secondary, "000000"),
        accent: safe(colors.accent, "CCCCCC"),
        background: safe(colors.background, "FFFFFF"),
        text: safe(colors.text, "000000"),
        footer: safe(colors.footer, "FFFFFF")
    };

    const safeFonts = {
        heading: safe(fonts.heading, "Arial"),
        body: safe(fonts.body, "Arial")
    };

    const safeStyles = {
        title: {
            fontSize: safeNum(styles.title?.fontSize, 32),
            bold: safe(styles.title?.bold, true)
        },
        body: {
            fontSize: safeNum(styles.body?.fontSize, 18)
        }
    };

    const safeLayout = {
        marginX: safeNum(layout.marginX, 1),
        marginY: safeNum(layout.marginY, 1),
        headerHeight: safeNum(layout.headerHeight, 1),
        footerHeight: safeNum(layout.footerHeight, 0.5)
    };

    const safeBg = {
        type: safe(bg.type, "solid"),
        color: safe(bg.color, safeColors.background)
    };

    // -------------------------
    // Metadata
    // -------------------------

    pptx.author = "Presentation Engine";
    pptx.company = safe(config.brand?.name, "Unknown Brand");
    pptx.subject = `${safe(config.brand?.name, "Brand")} Presentation`;
    pptx.title = `${safe(config.brand?.name, "Brand")} Slides`;

    // -------------------------
    // Slide Masters
    // -------------------------

    function defineSlideMasters() {

        // TITLE MASTER
        pptx.defineSlideMaster({
            title: "TITLE_MASTER",
            background: {color: safeBg.color},
            objects: [
                {
                    rect: {
                        x: 0,
                        y: 0,
                        w: 13.33,
                        h: 7.5,
                        fill: {color: safeColors.primary}
                    }
                },
                {
                    rect: {
                        x: 1,
                        y: 6.8,
                        w: 11.33,
                        h: 0.3,
                        fill: {color: safeColors.secondary}
                    }
                },
                {
                    placeholder: {
                        options: {
                            name: "title",
                            x: 1.5,
                            y: 2.5,
                            w: 10.33,
                            h: 1.2,
                            fontSize: safeStyles.title.fontSize,
                            fontFace: safeFonts.heading,
                            bold: safeStyles.title.bold,
                            color: safeColors.background,
                            align: "center",
                            valign: "middle"
                        },
                        text: "Click to add title"
                    }
                },
                {
                    placeholder: {
                        options: {
                            name: "subtitle",
                            x: 2,
                            y: 4,
                            w: 9.33,
                            h: 0.8,
                            fontSize: safeStyles.body.fontSize + 2,
                            fontFace: safeFonts.body,
                            color: safeColors.background,
                            align: "center",
                            valign: "middle"
                        },
                        text: "Click to add subtitle"
                    }
                }
            ]
        });

        // CONTENT MASTER
        pptx.defineSlideMaster({
            title: "CONTENT_MASTER",
            background: {color: safeBg.color},
            objects: [
                {
                    rect: {
                        x: 0,
                        y: 0,
                        w: 13.33,
                        h: safeLayout.headerHeight,
                        fill: {color: safeColors.background}
                    }
                },
                {
                    rect: {
                        x: 0,
                        y: 7.5 - safeLayout.footerHeight,
                        w: 13.33,
                        h: safeLayout.footerHeight,
                        fill: {color: safeColors.footer}
                    }
                },
                {
                    placeholder: {
                        options: {
                            name: "title",
                            x: safeLayout.marginX + 1,
                            y: 1.2,
                            w: 8.5,
                            h: 0.8,
                            fontSize: safeStyles.title.fontSize,
                            fontFace: safeFonts.heading,
                            bold: safeStyles.title.bold,
                            color: safeColors.primary
                        },
                        text: "Click to add title"
                    }
                },
                {
                    placeholder: {
                        options: {
                            name: "body",
                            x: safeLayout.marginX + 0.2,
                            y: 2.3,
                            w: 13.33 - (2 * safeLayout.marginX) - 0.4,
                            h: 7.5 - 2.3 - safeLayout.footerHeight - 0.2,
                            fontSize: safeStyles.body.fontSize,
                            fontFace: safeFonts.body,
                            color: safeColors.text
                        },
                        text: "Click to add content"
                    }
                }
            ]
        });

        // SECTION MASTER
        pptx.defineSlideMaster({
            title: "SECTION_MASTER",
            background: {color: safeBg.color},
            objects: [
                {
                    rect: {
                        x: 0,
                        y: 0,
                        w: 13.33,
                        h: 7.5,
                        fill: {color: safeColors.primary}
                    }
                },
                {
                    rect: {
                        x: 0,
                        y: 0,
                        w: 3,
                        h: 7.5,
                        fill: {color: safeColors.secondary}
                    }
                },
                {
                    placeholder: {
                        options: {
                            name: "title",
                            x: 3.5,
                            y: 2,
                            w: 8,
                            h: 1.5,
                            fontSize: safeStyles.title.fontSize + 8,
                            fontFace: safeFonts.heading,
                            bold: true,
                            color: safeColors.background
                        },
                        text: "Click to add section title"
                    }
                }
            ]
        });
    }

    defineSlideMasters();

    // -------------------------
    // Slides
    // -------------------------

    pptx.addSlide({masterName: "TITLE_MASTER"});
    pptx.addSlide({masterName: "CONTENT_MASTER"});
    pptx.addSlide({masterName: "SECTION_MASTER"});
    pptx.addSlide({masterName: "CONTENT_MASTER"});

    await pptx.writeFile({fileName: outputFile});

    return {
        success: true,
        output_file: outputFile
    };
}

// -------------------------
// Main execution
// -------------------------

(async () => {
    try {
        const config = await readStdin();
        const result = await buildPresentation(config);

        process.stdout.write(JSON.stringify({
            success: true,
            output_file: result.output_file
        }));

    } catch (error) {
        process.stdout.write(JSON.stringify({
            success: false,
            error: error.message
        }));

        process.exit(1);
    }
})();