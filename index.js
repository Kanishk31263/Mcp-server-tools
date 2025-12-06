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
// Markdown template
const MARKDOWN_TEMPLATE = `---
title: Presentation Title
topic: Main Topic
type: Corporate Presentation
author: Your Name
department: Department Name
brandPrimary: "#003366"
brandSecondary: "#00AEEF"
logo: "./logo.png"
---

## [plan] Agenda
- Introduction
- Key Problems
- Proposed Solution
- Architecture
- Benefits
- Summary

## [divider] 🔹 Section 1: Introduction

## [content] What is {{topic}}?
- Point 1
- Point 2
- Point 3

## [quote] Key Insight
> “A simple quote or important statement related to the topic.”

## [content] Why It Matters
- Business impact
- Technical relevance
- Real-world examples

## [divider] 🔹 Section 2: Problem Statement

## [content] Current Challenges
- Challenge 1
- Challenge 2
- Challenge 3

## [chart] Sample Bar Chart
type: bar
title: "Example Data"
xAxis: ["Q1","Q2","Q3","Q4"]
data:
  - label: "Revenue"
    values: [10, 15, 20, 25]

## [divider] 🔹 Section 3: Proposed Solution

## [content] Our Approach
- Strategy 1
- Strategy 2
- Strategy 3


## [content] Architecture Overview
- Component A
- Component B
- Component C

## [divider] 🔹 Section 4: Benefits

## [content] Why This Solution Works
- Benefit 1
- Benefit 2
- Benefit 3

## [quote] Key Takeaway
> “Short powerful summary line.”

## [divider] 🔹 Section 5: Summary

## [content] Summary of the Presentation
- Revisit main points
- Final notes
- Next steps

## [content] Thank You!
- Contact: your.email@company.com
- Department: Your Department
`;

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