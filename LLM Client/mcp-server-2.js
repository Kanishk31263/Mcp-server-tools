import express from "express";
import bodyParser from "body-parser";
import dotenv from "dotenv";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import path from 'path'
import { init } from "@heyputer/puter.js/src/init.cjs";

import { fileURLToPath } from 'url';
import { WebSocketServer } from "ws";

dotenv.config();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Build path to pptGeneratorServer/index.js
const MCP_server_path = path.join(__dirname, "../pptGeneratorServer/index.js");
const puter = init(process.env.PUPPER_AI_TOKEN);

const app = express();
const PORT =  3000;

// Middleware to parse JSON
app.use(bodyParser.json());

const mcpclient = new Client({ name: 'ppt-client', version: '1.0.0' });

const templateMarkDown = `---
discipline: Pop Culture / Film Studies
type: Lesson
module: Alien Franchise
lesson: Xenomorph Biology & Lore
---

# Xenomorphs: The Ultimate Alien

## [title] Introduction
- Xenomorph is the iconic extraterrestrial creature from the *Alien* movie series.
- First appeared in *Alien* (1979) directed by Ridley Scott.
- Known for its terrifying life cycle and adaptability.

## [bullet] Physical Characteristics
- Exoskeleton-like, black or dark grey body.
- Elongated head with no visible eyes.
- Inner retractable jaws for attacking prey.
- Highly agile and strong, capable of climbing walls and ceilings.

## [bullet] Life Cycle
1. Egg (Laid by Queen)
2. Facehugger (attaches to host)
3. Chestburster (emerges violently from host)
4. Adult Xenomorph (fully grown, lethal predator)

## [image] Xenomorph Lifecycle
- Optional slide with diagram showing Egg → Facehugger → Chestburster → Adult
- Image path example: assets/xenomorph-lifecycle.png

## [bullet] Behavior & Abilities
- Highly intelligent and adaptive predator.
- Can survive in extreme environments, including outer space.
- Uses stealth and ambush tactics.
- Acidic blood as defense mechanism.

## [bullet] Queen Xenomorph
- Larger, more intelligent variant
- Lays eggs for reproduction
- Commands drones in hive-like structures

## [bullet] Cultural Impact
- Influenced countless sci-fi and horror films.
- Recognized for terrifying design by H.R. Giger.
- Spawned video games, comics, and merchandise.

## [bullet] Trivia
- The chestburster scene in *Alien* (1979) is one of the most iconic horror moments in cinema.
- The Xenomorph’s design is biomechanical, blending organic and mechanical features.
- No two Xenomorphs look exactly alike due to host-based adaptation.

## [title] Conclusion
- Xenomorphs are the ultimate embodiment of sci-fi horror.
- Represent a perfect predator with a complex life cycle.
- Continue to influence pop culture decades after their creation.

`

const transport = new StdioClientTransport({
  command: process.execPath,
  args: [path.resolve(MCP_server_path)]
});
await mcpclient.connect(transport);


// -------------------- Start server --------------------
const server = app.listen(PORT, () => {
  console.log(`Express server running on http://localhost:${PORT}`);
});
// Create WebSocket server attached to Express
const wss = new WebSocketServer({ server });

wss.on("connection", (socket) => {
  console.log("Client connected via WebSocket");

  socket.on("message", async (msg) => {
    try {
      const { topic } = JSON.parse(msg.toString());

      // SEND ACK
      socket.send(JSON.stringify({ status: "processing", topic }));

      // CALL PUTER AI
      const response  = await puter.ai.chat(`
Return ONLY a valid JSON object in this exact structure:

{
  "filename": "<topic_based_filename>.pptx",
  "markdown": "<proper_markdown_here>"
}

CRITICAL RULES:
- Return ONLY JSON. No backticks. No explanations. No extra text.
- "markdown" must follow the EXACT template format below.
- Slides must use ONLY these types:

  ## [title] Title text
  ## [bullet] Bullet section title
  ## [image] Image section title
  ## [divider] Section divider title

- Divider slides must look like:

  ## [divider] Section Name

- Do NOT add “Slide 1”, “Slide 2”, etc.
- Do NOT add numbering in slide titles.
- Use "---" ONLY if it already exists in templateMarkDown (otherwise avoid).

Here is the slide template you must follow exactly:
${templateMarkDown}

Generate content for the topic: "${topic}"
`);

      const raw = response.message.content;
      const parsed = JSON.parse(raw);

      // CALL MCP TOOL
      await mcpclient.callTool({
        name: "generate_presentation",
        arguments: {
          markdown: parsed.markdown,
          filename: parsed.filename
        }
      });

      // SEND RESULT BACK
      socket.send(JSON.stringify({
        status: "done",
        filename: parsed.filename
      }));

    } catch (error) {
      console.error("WS Error:", error);
      socket.send(JSON.stringify({ status: "error", error: error.message }));
    }
  });
});

