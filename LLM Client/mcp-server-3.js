import express from "express";
import multer from "multer";
import dotenv from "dotenv";
import { WebSocketServer } from "ws";
import bodyParser from "body-parser";

import path from "path";
import { fileURLToPath } from 'url';
import cors from 'cors';
import { PDFParse } from 'pdf-parse';

dotenv.config();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Build path to pptGeneratorServer/index.js
const MCP_server_path = path.join(__dirname, "../pptGeneratorServer/index.js");
const app = express();
app.use(cors({
    origin: "*",
    methods: ["GET", "POST"],
    allowedHeaders: ["Content-Type"]
}));

app.use(bodyParser.json());
const upload = multer({ dest: "uploads/" });

// =============== WEBSOCKET ===============
let uiClient = null;
const wss = new WebSocketServer({ port: 3001 });

wss.on("connection", (socket) => {
    uiClient = socket;
    socket.send("Connected to PPT Generator");
});

function status(msg) {
    if (uiClient) uiClient.send(msg);
}

app.post("/upload-pdf", upload.single("resume"), async (req, res) => {
    try {
        const pdfPath = req.file.path;

        status("PDF received… Extracting text...");
        const parser = new PDFParse({ url:pdfPath });        
        let  result = await parser.getText();
        await parser.destroy();
        console.log('text data',result);

        // status("PDF text extracted. Sending to LLM...");
        // const markdown = await generateMarkdownFromAI(text);

        status("Markdown ready. Calling MCP tool...");
        const filename = "Generated_Presentation.pptx";


        status("✅ PPT Successfully Generated!");
        res.json({ ok: true, filename });
    } catch (e) {
        console.error(e);
        status("❌ Error: " + e.message);
        res.status(500).json({ error: e.message });
    }
});


function cleanPDFText(text) {
    return text
        .replace(/--\s*\d+\s*of\s*\d+\s*--/gi, '')
        .replace(/Page\s+\d+\s+of\s+\d+/gi, '')
        .split('\n')
        .filter(line => line.trim() !== '')
        .map(line => {
            if (line.match(/\b[A-Z]\s+[A-Z]/)) {
                return line.replace(/\s+/g, '');
            }
            return line;
        })
        .join('\n')
        .replace(/\s+/g, ' ')
        .replace(/\n\s*\n+/g, '\n\n')
        .trim();
}


app.listen(3000, () => console.log("Server running on http://localhost:3000"));