// MCP server that exposes the presentation-engine scripts as tools over HTTP.
// Agents call POST /mcp with JSON-RPC; the server spawns the relevant Node script
// as a child process and returns the result. No script is modified.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import express from "express";
import { spawnSync } from "child_process";
import { writeFileSync, unlinkSync, readFileSync } from "fs";
import { join } from "path";
import { tmpdir } from "os";
import { fileURLToPath } from "url";
import { dirname, resolve } from "path";

// Resolves the scripts path relative to this file, so it works inside Docker too
const __dirname = dirname(fileURLToPath(import.meta.url));
const SCRIPTS_DIR = resolve(__dirname, "../../presentation-engine/src");

// Runs a Node script synchronously and normalises the result into {success, stdout, stderr}
function runScript(scriptPath, args) {
  const result = spawnSync("node", [scriptPath, ...args], { encoding: "utf8" });
  return {
    success: result.status === 0,
    stdout: result.stdout?.trim(),
    stderr: result.stderr?.trim()
  };
}

// The scripts expect file paths, not raw JSON — this bridges the gap via OS temp dir
function writeTmp(name, content) {
  const path = join(tmpdir(), name);
  writeFileSync(path, JSON.stringify(content, null, 2), "utf8");
  return path;
}

// A new McpServer instance is created per request (see app.post below) to keep
// each call fully stateless, which is required for serverless platforms like Code Engine
function createServer() {
  const server = new McpServer({ name: "pptx-tools", version: "1.0.0" });

  // Tool 1: validates the presentation content JSON (used by the Content Agent)
  server.registerTool(
    "validate_presentation",
    {
      description: "Validates a presentation JSON against the smart_presentation_v2 schema",
      inputSchema: {
         presentation: z.object({}).passthrough() 
      },
    },
    async ({ presentation }) => {
      const tmpPath = writeTmp(`presentation_${Date.now()}.json`, presentation);
      const result = runScript(join(SCRIPTS_DIR, "validate_presentation.js"), [tmpPath]);
      unlinkSync(tmpPath);
      return {
        content: [{ type: "text", text: result.success ? result.stdout : result.stderr }],
        isError: !result.success
      };
    }
  );

  // Tool 2: validates the company style profile JSON (used by the Format Agent)
  server.registerTool(
    "validate_profile",
    {
      description: "Validates a company style profile JSON against the company_style_profile_v1 schema",
      inputSchema: {
        profile: z.object({}).passthrough()
      },
    },
    async ({ profile }) => {
      const tmpPath = writeTmp(`profile_${Date.now()}.json`, profile);
      const result = runScript(join(SCRIPTS_DIR, "validate_profile.js"), [tmpPath]);
      unlinkSync(tmpPath);
      return {
        content: [{ type: "text", text: result.success ? result.stdout : result.stderr }],
        isError: !result.success
      };
    }
  );

  // Tool 3: renders the final PPTX (used by the PPTX Agent after both validations pass).
  // Returns the file as base64 because Code Engine has no persistent filesystem.
  server.registerTool(
    "render_presentation",
    {
      description:"Renders a validated presentation JSON (with optional style profile) into a PPTX file. Returns the file as base64.",
      inputSchema: {
        presentation: z.object({}).passthrough(),
        output_filename: z.string().default("output.pptx"),
        profile: z.object({}).passthrough().optional()
      },
    },
    async ({ presentation, output_filename, profile }) => {
      const suffix = Date.now();
      const contentPath = writeTmp(`presentation_${suffix}.json`, presentation);
      const outputPath = join(tmpdir(), `${suffix}_${output_filename}`);
      const args = [contentPath, outputPath];

      let profilePath = null;
      if (profile) {
        profilePath = writeTmp(`profile_${suffix}.json`, profile);
        args.push(profilePath);
      }

      const result = runScript(join(SCRIPTS_DIR, "render_presentation.js"), args);

      // Clean up inputs regardless of outcome
      unlinkSync(contentPath);
      if (profilePath) unlinkSync(profilePath);

      if (!result.success) {
        return {
          content: [{ type: "text", text: result.stderr }],
          isError: true
        };
      }

      const base64 = readFileSync(outputPath).toString("base64");
      unlinkSync(outputPath);

      return {
        content: [
          { type: "text", text: `PPTX rendered successfully. Filename: ${output_filename}` },
          { type: "text", text: base64 }
        ]
      };
    }
  );

   // Tool 4: validates the generated brand validation rules JSON
  server.registerTool(
    "validate_rules",
    {
      description: "Validates a brand validation rules JSON against the brand_validation_rules_v1 schema",
      inputSchema: {
        rules: z.object({}).passthrough()
      },
    },
    async ({ rules }) => {
      const tmpPath = writeTmp(`rules_${Date.now()}.json`, rules);
      const result = runScript(join(SCRIPTS_DIR, "validate_rules.js"), [tmpPath]);
      unlinkSync(tmpPath);

      return {
        content: [{ type: "text", text: result.success ? result.stdout : result.stderr }],
        isError: !result.success
      };
    }
  );

   // Tool 5: validates a presentation JSON against generated brand validation rules
  server.registerTool(
    "validate_deck_against_rules",
    {
      description: "Checks whether a presentation JSON follows the generated brand validation rules. Returns a structured validation report.",
      inputSchema: {
        presentation: z.object({}).passthrough(),
        rules: z.object({}).passthrough()
      },
    },
    async ({ presentation, rules }) => {
      const suffix = Date.now();
      const presentationPath = writeTmp(`presentation_${suffix}.json`, presentation);
      const rulesPath = writeTmp(`rules_${suffix}.json`, rules);

      const result = runScript(
        join(SCRIPTS_DIR, "validate_deck_against_rules.js"),
        [presentationPath, rulesPath]
      );

      unlinkSync(presentationPath);
      unlinkSync(rulesPath);

      return {
        content: [{ type: "text", text: result.stdout || result.stderr }],
        isError: !result.success
      };
    }
  );

  // Tool 6: generates a branded PPTX from a presentation brief and brand guidelines
  server.registerTool(
    "generate_presentation",
    {
      description: "Generates a branded PowerPoint presentation from a presentation brief and brand guidelines. Returns the file as base64.",
      inputSchema: {
        json: z.object({}).passthrough(),
        output_filename: z.string().default("presentation.pptx"),
      },
    },
    async ({ json, output_filename }) => {
      const suffix = Date.now();
      const inputPath = writeTmp(`giovanni_input_${suffix}.json`, json);
      const outputPath = join(tmpdir(), `${suffix}_${output_filename}`);

      const result = runScript(join(SCRIPTS_DIR, "render_giovanni.js"), [inputPath, outputPath]);
      unlinkSync(inputPath);

      if (!result.success) {
        return { content: [{ type: "text", text: result.stderr }], isError: true };
      }

      const base64 = readFileSync(outputPath).toString("base64");
      unlinkSync(outputPath);

      return {
        content: [
          { type: "text", text: `Presentation generated successfully. Filename: ${output_filename}` },
          { type: "text", text: base64 }
        ]
      };
    }
  );

  return server;
}

const app = express();
app.use(express.json());

// Each POST creates a fresh server+transport pair — stateless by design
app.post("/mcp", async (req, res) => {
  const server = createServer();
  // sessionIdGenerator: undefined disables session tracking (correct for stateless deployments)
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
  res.on("finish", () => server.close());
});

// Required by IBM Code Engine for health/readiness checks
app.get("/health", (_, res) => res.json({ status: "ok" }));

app.listen(8080, () => console.error("pptx-mcp-server running on :8080"));
