# Agentic AI Challenge

This repository contains our solution for the IBM Watsonx Agentic AI Challenge 2026.

---

An agentic system for **PowerPoint generation, restyling, and analysis**. Users interact with natural-language agents (built on IBM watsonx Orchestrate) that use Python tools and a Node.js rendering engine to produce branded `.pptx` files end-to-end.

---

## What it does

Three capabilities, each exposed through its own agent:

| Agent | Purpose |
|------|--------|
| [siro_simple_agent](agents/siro_simple_agent.yaml) | Generates a new PowerPoint from a natural-language brief |
| [siro_simple_agent_2](agents/siro_simple_agent_2.yaml) | Extracts an existing PowerPoint's content and re-renders it in a clean, brand-consistent style |
| [sarinas_format_agent](agents/sarinas_format_agent.yaml) | Analyzes a PowerPoint and reports its formatting (colors, fonts, fills, borders, layout) |

---

## Architecture

```
User ─▶ Agent (watsonx Orchestrate, LLM = groq/openai/gpt-oss-120b)
          │
          ├─▶ Python tools                  (extraction / format inspection)
          │     • siro_extract_pptx
          │     • pptx_format_tool (extract_pptx_formatting)
          │
          └─▶ MCP server (HTTP, Node.js)    (rendering)
                • siro_simple_render        ─▶ presentation-engine (pptxgenjs)
                                                   │
                                                   └─▶ .pptx file + public download URL
```

The agent decides which tools to call. The MCP server is a thin HTTP wrapper that spawns the Node renderer scripts and returns a download URL for the generated file.

---

## Repository layout

```
agents/                     watsonx Orchestrate agent definitions (YAML)
  siro_simple_agent.yaml         create from prompt
  siro_simple_agent_2.yaml       extract + restyle uploaded .pptx
  sarinas_format_agent.yaml      inspect formatting of a .pptx

tools/                      Python tools registered in watsonx Orchestrate
  siro_extract_pptx.py           reads .pptx, returns structured slide content
  pptx_format_tool.py            reads .pptx, returns colors/fonts/fills/borders
  requirements.txt

services/mcp-pptx/          MCP HTTP server (Node.js) — rendering bridge
  index.js                       JSON-RPC over /mcp, public /files endpoint
  Dockerfile                     deploys on IBM Code Engine

presentation-engine/        JSON → .pptx renderer (pptxgenjs)
  src/siro_simple_render.js      simple-format renderer used by the agents
  src/render_giovanni.js         richer branded renderer
  readme.md                      engine-level docs

knowledge/                  Brand specifications
  specs.json                     Talentia color tokens & logo references

assets/talentia/            Talentia brand assets (logos, asset-pack, policy)
```

---

## How a request flows

**Generating a deck from a prompt** ([siro_simple_agent](agents/siro_simple_agent.yaml)):

1. User describes the deck in natural language.
2. The agent infers topic, audience, slide count, language, tone, and style defaults.
3. It builds a structured presentation JSON (`title`, `subtitle`, `settings`, `slides[]` with types `title` / `bullets` / `two_column` / `closing`).
4. It calls `mcp-pptx-v2:siro_simple_render` over MCP.
5. The MCP server runs [siro_simple_render.js](presentation-engine/src/siro_simple_render.js), saves the file to its public download folder, and returns a `download_url`.
6. The agent replies with the link.

**Restyling an uploaded deck** ([siro_simple_agent_2](agents/siro_simple_agent_2.yaml)):

1. User uploads a `.pptx`.
2. Agent calls `siro_extract_pptx` (Python) → structured slide content.
3. Agent maps content to the simple slide schema.
4. Agent calls `siro_simple_render` → branded output → download URL.

**Inspecting formatting** ([sarinas_format_agent](agents/sarinas_format_agent.yaml)):

1. User uploads a `.pptx`.
2. Agent calls `extract_pptx_formatting` (Python) → JSON of colors, fonts, fills, borders, etc.
3. Agent presents the result, optionally filtered by the user's question.

---

## Slide schema (simple format)

The renderer accepts a single JSON document:

```json
{
  "title": "Deck title",
  "subtitle": "Optional subtitle",
  "settings": {
    "fontFace": "Arial",
    "titleColor": "24135F",
    "bodyColor": "1D1D1D",
    "backgroundColor": "FFFFFF",
    "accentColor": "D0006F",
    "footerText": "Generated presentation"
  },
  "slides": [
    { "type": "title",     "title": "...", "subtitle": "..." },
    { "type": "bullets",   "title": "...", "bullets": ["...", "..."] },
    { "type": "two_column","title": "...",
      "leftTitle": "...",  "leftBullets":  ["..."],
      "rightTitle": "...", "rightBullets": ["..."] },
    { "type": "closing",   "title": "...", "bullets": ["..."] }
  ]
}
```

Colors are hex strings without `#`. Default brand palette is Talentia (see [knowledge/specs.json](knowledge/specs.json)).

---

## Running locally

### Renderer (Node.js)

```bash
cd presentation-engine
npm install
node src/siro_simple_render.js <input.json> <output.pptx>
```

### MCP server (Docker)

```bash
docker build -t mcp-pptx -f services/mcp-pptx/Dockerfile .
docker run -p 8080:8080 -e PUBLIC_BASE_URL=http://localhost:8080 mcp-pptx
```

- `POST /mcp` — JSON-RPC entry point for agents
- `GET  /files/<filename>` — public download for generated PPTX
- `GET  /health` — readiness probe (used by IBM Code Engine)

`PUBLIC_BASE_URL` must point to the externally reachable host; without it the agent can render but cannot return a download link.

### Python tools

```bash
cd tools
pip install -r requirements.txt
```

The tools depend on `ibm_watsonx_orchestrate.agent_builder.tools` and are registered in watsonx Orchestrate, not run standalone.

---

## Deployment

The MCP server is built for **IBM Code Engine** (stateless, port 8080, `/health` probe). The agents run on **IBM watsonx Orchestrate** with `groq/openai/gpt-oss-120b` as the underlying LLM. Python tools are uploaded to Orchestrate and bound to the agents by name; the MCP tool is bound under the namespace `mcp-pptx-v2`.

---

## Brand assets

[assets/talentia/](assets/talentia/) holds the Talentia logo set and a brand policy. [knowledge/specs.json](knowledge/specs.json) is the canonical source for color tokens and logo references used by the renderer's default profile.
