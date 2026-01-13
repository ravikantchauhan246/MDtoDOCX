# MD to DOCX Converter

A powerful web application that converts Markdown files to professionally formatted DOCX (Microsoft Word) documents with AI-powered diagram conversion.

## Features

- **GitHub-Style Preview** - Real-time markdown preview with GitHub markdown styling
- **DOCX Generation** - Convert markdown to Word documents with proper formatting
- **AI-Powered Parsing** - Uses Google Gemini AI to intelligently parse entire documents
- **Smart Diagram Conversion** - Automatically converts ASCII diagrams to structured, readable components
- **Preserves Formatting** - Maintains:
  - Code blocks with syntax highlighting
  - Tables with proper borders and spacing
  - Headings (H1-H6)
  - Lists (ordered and unordered)
  - Blockquotes
  - Inline formatting (bold, italic, code, links)
  - Emojis
  - Horizontal rules
- **Fallback Mode** - Works even without AI (uses local parsing)
- **Multiple Input Methods** - Upload files or paste markdown directly
- **Easy Download** - One-click download of converted DOCX files

## Quick Start

### Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- Google Gemini API key (optional but recommended for best results)

### Installation

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd MDtoDOCXNode
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Set up environment variables**
   
   Create a `.env` file in the project root:
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` and add your Gemini API key:
   ```env
   GEMINI_API_KEY=your_api_key_here
   PORT=3000
   ```

   > **Get a free Gemini API key:** Visit [Google AI Studio](https://makersuite.google.com/app/apikey)

4. **Start the server**
   ```bash
   npm start
   ```

5. **Open your browser**
   
   Navigate to `http://localhost:3000`

## Usage

### Web Interface

1. **Upload a File** - Drag and drop a `.md` file or click to browse
2. **Paste Markdown** - Type or paste markdown content directly
3. **Preview** - See a GitHub-style preview of your document
4. **Toggle AI** - Enable/disable AI-powered parsing (checkbox in UI)
5. **Download** - Click "Download DOCX" to get your Word document

### API Endpoints

#### Preview Markdown
```bash
POST /api/preview
Content-Type: application/json

{
  "markdown": "# Hello World\nYour markdown here..."
}
```

#### Convert to DOCX
```bash
POST /api/convert
Content-Type: application/json

{
  "markdown": "# Document Title\n...",
  "filename": "output.md",
  "useAI": true
}
```

#### Upload File
```bash
POST /api/upload
Content-Type: multipart/form-data

file: <markdown-file>
```

#### Check Gemini Status
```bash
GET /api/gemini/status
```

## Project Structure

```
MDtoDOCXNode/
├── public/                 # Frontend assets
│   ├── index.html         # Main web interface
│   ├── css/
│   │   ├── styles.css     # Custom styles
│   │   └── github-markdown.css  # GitHub markdown styling
│   └── js/
│       └── app.js         # Frontend logic
├── src/                   # Backend source code
│   ├── server.js          # Express server
│   ├── parser/
│   │   └── markdownParser.js  # Markdown parsing logic
│   ├── converter/
│   │   └── docxConverter.js   # DOCX generation
│   └── services/
│       └── geminiService.js   # Google Gemini API integration
├── .env                   # Environment variables (create this)
├── .env.example           # Environment template
├── package.json           # Dependencies and scripts
└── README.md             # This file
```

## Configuration

### Font Sizes

Default font sizes (in points):
- **Headings:** H1: 28, H2: 24, H3: 20, H4: 18, H5: 16, H6: 14
- **Body:** 11pt
- **Code:** 10pt
- **Tables:** 10pt

Edit `src/converter/docxConverter.js` to customize.

### Gemini API Settings

**Retry Configuration:**
- Max retries: 2
- Base delay: 2 seconds
- Max delay: 20 seconds

**Model:** `gemini-2.0-flash`

Edit `src/services/geminiService.js` to customize.

## How AI Parsing Works

1. **Full Document Parsing** - Entire markdown is sent to Gemini in one API call
2. **Structured Analysis** - Gemini returns JSON with typed elements:
   - Headings with levels
   - Paragraphs with inline formatting
   - Code blocks with language tags
   - Tables with headers and rows
   - Lists (ordered/unordered)
   - Diagrams with components and connections
3. **DOCX Generation** - Structured data is converted to formatted Word elements
4. **Fallback** - If AI fails (rate limit/error), uses local token-based parsing

## Diagram Conversion

ASCII diagrams are automatically detected and converted to rich components:

**Input:**
```
┌─────────────┐
│   Client    │
└──────┬──────┘
       │
       ↓
┌─────────────┐
│   Server    │
└─────────────┘
```

**Output in DOCX:**
- Title and description
- Component table with names and descriptions
- Data flow connections with arrows
- Styled with colored backgrounds and borders

## Technologies Used

- **Backend:**
  - Node.js + Express.js
  - markdown-it (Markdown parser)
  - docx (DOCX generation)
  - @google/generative-ai (Gemini AI)
  - highlight.js (Syntax highlighting)
  - multer (File uploads)
  - dotenv (Environment config)

- **Frontend:**
  - Vanilla JavaScript
  - GitHub Markdown CSS
  - Drag-and-drop API

## Rate Limiting

Gemini free tier has rate limits. The app includes:
- **Exponential backoff** with jitter
- **Suggested delay extraction** from error messages
- **Automatic retry** (up to 2 retries)
- **Graceful fallback** to local parsing

## Troubleshooting

### "Gemini API not initialized"
- Check your `.env` file has `GEMINI_API_KEY`
- Verify the API key is valid
- The app still works without AI (local parsing mode)

### "429 Too Many Requests"
- You've hit Gemini's rate limit
- Wait a few seconds and try again
- Disable AI toggle to use local parsing
- Consider upgrading to paid tier for higher limits

### Tables appear cramped
- Font sizes have been optimized (11pt body, 10pt tables)
- Cell margins are set to 100/150 twips
- Tables use autofit layout
- Adjust `STYLES` in `docxConverter.js` if needed

### Diagrams not converting properly
- Enable AI toggle for best diagram conversion
- Ensure diagram uses standard ASCII box-drawing characters
- Fallback mode will display as code block

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues or questions, please open an issue on the repository.

---

**Made with Node.js, Express, and Google Gemini AI**
