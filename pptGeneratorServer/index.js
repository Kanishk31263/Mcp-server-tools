#!/usr/bin/env node
/**
 * VITI PPTX Generator - MCP Server
 * 
 * Generates lesson presentations from markdown content.
 * 
 * Tools:
 *   - generate_presentation: Create PPTX from markdown
 *   - get_template: Get markdown template example
 *   - get_config: View current configuration
 *   - update_instructor: Update instructor info
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { generatePresentation } from './ppt-generator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// Create server instance
const server = new Server(
  { name: 'pptx-generator2', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

// Define available tools
server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'generate_presentation',
      description: 'Generate a PPTX presentation from markdown content. Returns path to the generated file.',
      inputSchema: {
        type: 'object',
        properties: {
          markdown: {
            type: 'string',
            description: 'Full markdown content including frontmatter (---) with discipline, type, module, lesson fields and slide sections marked with ## [type] Title'
          },
          filename: {
            type: 'string',
            description: 'Output filename without extension (e.g., "python-lesson-5")'
          }
        },
        required: ['markdown', 'filename']
      }
    }
  ]
}));

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  console.error('args',args,'name',name)
  
  try {
    switch (name) {
      case 'generate_presentation': {
        const filename = args.filename.replace(/\.pptx$/i, '');
        const outputPath = path.join(OUTPUT_DIR, `${filename}.pptx`);
        
        await generatePresentation(args.markdown, outputPath, __dirname);
        
        return {
          content: [{
            type: 'text',
            text: `✅PPt generated at  ${outputPath}.`
          }]
        };
      }
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `❌ : ${error.message}`
      }],
      isError: true
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("pptx generater 2  Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});