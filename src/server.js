import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { parseMarkdown } from './parser/markdownParser.js';
import { convertToDocx } from './converter/docxConverter.js';
import { initializeGemini, disableGemini, enableGemini, isGeminiAvailable } from './services/geminiService.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize Gemini API
initializeGemini(process.env.GEMINI_API_KEY);

// Middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname, '../public')));

// Multer configuration for file uploads
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (ext === '.md' || ext === '.markdown' || ext === '.txt') {
            cb(null, true);
        } else {
            cb(new Error('Only markdown files (.md, .markdown, .txt) are allowed'));
        }
    }
});

// API Routes

// Preview markdown as HTML
app.post('/api/preview', (req, res) => {
    try {
        const { markdown } = req.body;
        if (!markdown) {
            return res.status(400).json({ error: 'No markdown content provided' });
        }
        const html = parseMarkdown(markdown);
        res.json({ html });
    } catch (error) {
        console.error('Preview error:', error);
        res.status(500).json({ error: 'Failed to parse markdown' });
    }
});

// Upload markdown file and get preview
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        const markdown = req.file.buffer.toString('utf-8');
        const html = parseMarkdown(markdown);
        res.json({ 
            html, 
            markdown,
            filename: req.file.originalname 
        });
    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: 'Failed to process uploaded file' });
    }
});

// Convert markdown to DOCX
app.post('/api/convert', async (req, res) => {
    try {
        const { markdown, filename, useAI } = req.body;
        if (!markdown) {
            return res.status(400).json({ error: 'No markdown content provided' });
        }
        
        // Convert to boolean
        const useAIFlag = useAI !== false;
        
        const docxBuffer = await convertToDocx(markdown, useAIFlag);
        const outputFilename = filename ? 
            filename.replace(/\.(md|markdown|txt)$/i, '.docx') : 
            'document.docx';
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
        res.send(Buffer.from(docxBuffer));
    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: 'Failed to convert to DOCX' });
    }
});

// Get Gemini status
app.get('/api/gemini/status', (req, res) => {
    res.json({ enabled: isGeminiAvailable() });
});

// Serve the main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../public/index.html'));
});

// Error handling middleware
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ error: 'File size exceeds 10MB limit' });
        }
    }
    console.error('Server error:', error);
    res.status(500).json({ error: error.message || 'Internal server error' });
});

app.listen(PORT, () => {
    console.log(`🚀 MD to DOCX Converter running at http://localhost:${PORT}`);
});
