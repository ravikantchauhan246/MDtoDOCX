import { GoogleGenerativeAI } from '@google/generative-ai';

// Initialize Gemini
let genAI = null;
let model = null;
let useGemini = true; // Can be disabled if rate limited

// Retry configuration
const RETRY_CONFIG = {
    maxRetries: 2,
    baseDelay: 2000, // 2 seconds
    maxDelay: 20000  // 20 seconds for rate limits
};

/**
 * Sleep for a given number of milliseconds
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Extract retry delay from error message if available
 */
function extractRetryDelay(error) {
    const message = error?.message || String(error);
    // Look for patterns like "retry in 17.369220798s" or "retryDelay":"17s"
    const match = message.match(/retry.*?(\d+\.?\d*)\s*s/i) || 
                  message.match(/retryDelay.*?(\d+)/i);
    if (match) {
        return Math.ceil(parseFloat(match[1]) * 1000) + 500; // Convert to ms + buffer
    }
    return null;
}

/**
 * Calculate exponential backoff delay
 */
function getRetryDelay(attempt, suggestedDelay = null) {
    if (suggestedDelay) {
        return Math.min(suggestedDelay, RETRY_CONFIG.maxDelay);
    }
    const delay = Math.min(
        RETRY_CONFIG.baseDelay * Math.pow(2, attempt),
        RETRY_CONFIG.maxDelay
    );
    // Add jitter (±20%)
    return delay * (0.8 + Math.random() * 0.4);
}

/**
 * Initialize the Gemini API client
 */
export function initializeGemini(apiKey) {
    if (!apiKey) {
        console.warn('⚠️ Gemini API key not provided. Diagram conversion will use fallback mode.');
        useGemini = false;
        return false;
    }
    try {
        genAI = new GoogleGenerativeAI(apiKey);
        model = genAI.getGenerativeModel({ model: 'gemini-3-flash-preview' });
        useGemini = true;
        console.log('✅ Gemini API initialized successfully');
        return true;
    } catch (error) {
        console.error('Failed to initialize Gemini:', error);
        useGemini = false;
        return false;
    }
}

/**
 * Disable Gemini (use fallback only)
 */
export function disableGemini() {
    useGemini = false;
    console.log('🔕 Gemini disabled, using fallback mode');
}

/**
 * Enable Gemini
 */
export function enableGemini() {
    if (model) {
        useGemini = true;
        console.log('🔔 Gemini enabled');
    }
}

/**
 * Check if Gemini is available and enabled
 */
export function isGeminiAvailable() {
    return model !== null && useGemini;
}

/**
 * Convert ASCII diagram to structured DOCX-friendly format using Gemini
 * Returns an object with structured data for DOCX generation
 * Includes retry logic for handling transient API errors (404, 503, etc.)
 */
export async function convertDiagramWithAI(asciiDiagram) {
    // Use fallback if Gemini is disabled or not available
    if (!model || !useGemini) {
        return fallbackDiagramConversion(asciiDiagram);
    }

    const prompt = `Analyze this ASCII diagram and convert it to a structured format for a Word document.

ASCII DIAGRAM:
\`\`\`
${asciiDiagram}
\`\`\`

Respond with a JSON object containing:
1. "type": One of "flowchart", "architecture", "hierarchy", "table", "sequence", "other"
2. "title": A short title for the diagram
3. "description": A brief description of what the diagram shows
4. "components": Array of components/boxes with their names and descriptions
5. "connections": Array of connections between components (from, to, label)
6. "summary": A well-formatted text summary that captures the diagram's meaning

IMPORTANT: Respond ONLY with valid JSON, no markdown code blocks.`;

    let lastError = null;
    
    for (let attempt = 0; attempt <= RETRY_CONFIG.maxRetries; attempt++) {
        try {
            if (attempt > 0) {
                // Check if we have a suggested delay from rate limit error
                const suggestedDelay = extractRetryDelay(lastError);
                const delay = getRetryDelay(attempt - 1, suggestedDelay);
                console.log(`🔄 Retry attempt ${attempt}/${RETRY_CONFIG.maxRetries} after ${Math.round(delay / 1000)}s...`);
                await sleep(delay);
            }
            
            const result = await model.generateContent(prompt);
            const response = await result.response;
            let text = response.text().trim();
            
            // Clean up response - remove markdown code blocks if present
            text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
            
            const parsed = JSON.parse(text);
            
            if (attempt > 0) {
                console.log(`✅ Gemini API succeeded on retry attempt ${attempt}`);
            }
            
            return {
                success: true,
                ...parsed
            };
        } catch (error) {
            lastError = error;
            const statusCode = error?.status || error?.response?.status || 'unknown';
            const isRetryable = statusCode === 404 || statusCode === 503 || statusCode === 429 || 
                               error.message?.includes('404') || error.message?.includes('503') ||
                               error.message?.includes('overloaded') || error.message?.includes('temporarily');
            
            console.error(`❌ Gemini API error (attempt ${attempt + 1}/${RETRY_CONFIG.maxRetries + 1}): ${error.message || error}`);
            
            if (!isRetryable || attempt >= RETRY_CONFIG.maxRetries) {
                console.error('Gemini diagram conversion failed after retries, using fallback');
                return fallbackDiagramConversion(asciiDiagram);
            }
        }
    }
    
    // Should not reach here, but just in case
    console.error('Gemini conversion exhausted retries:', lastError);
    return fallbackDiagramConversion(asciiDiagram);
}

/**
 * Fallback conversion when Gemini is not available
 */
function fallbackDiagramConversion(asciiDiagram) {
    // Extract text content from ASCII diagram
    const lines = asciiDiagram.split('\n');
    const textContent = [];
    
    for (const line of lines) {
        // Remove box-drawing characters and extract text
        const cleaned = line
            .replace(/[─│┌┐└┘├┤┬┴┼═║╔╗╚╝╠╣╦╩╬▼▲◄►●○■□▪▫\[\]]/g, '')
            .replace(/[-|+*]/g, '')
            .trim();
        
        if (cleaned.length > 2) {
            textContent.push(cleaned);
        }
    }
    
    // Deduplicate and filter
    const uniqueContent = [...new Set(textContent)].filter(t => t.length > 2);
    
    return {
        success: true,
        type: 'architecture',
        title: 'System Architecture Diagram',
        description: 'Diagram converted from ASCII representation',
        components: uniqueContent.slice(0, 20).map((text, i) => ({
            name: text,
            description: ''
        })),
        connections: [],
        summary: uniqueContent.join('\n• ')
    };
}

/**
 * Generate a DOCX-friendly representation of a diagram
 * Returns structured elements for the DOCX builder
 */
export async function processDiagramForDocx(asciiDiagram) {
    const analysis = await convertDiagramWithAI(asciiDiagram);
    
    return {
        type: 'diagram',
        data: analysis
    };
}

/**
 * Parse entire markdown document using Gemini
 * Returns structured document elements for DOCX generation
 */
export async function parseDocumentWithAI(markdown) {
    if (!model || !useGemini) {
        return null; // Return null to signal fallback should be used
    }

    const prompt = `You are a document parser. Parse this Markdown document and convert it to a structured JSON format for generating a Word document.

MARKDOWN DOCUMENT:
\`\`\`markdown
${markdown}
\`\`\`

Return a JSON object with this structure:
{
  "title": "Document title (from first h1 or inferred)",
  "elements": [
    {
      "type": "heading",
      "level": 1-6,
      "text": "Heading text (include emoji if present)"
    },
    {
      "type": "paragraph",
      "content": [
        { "text": "plain text" },
        { "text": "bold text", "bold": true },
        { "text": "italic text", "italic": true },
        { "text": "code", "code": true },
        { "text": "link text", "link": "url" }
      ]
    },
    {
      "type": "code",
      "language": "javascript",
      "content": "code here preserving all whitespace and newlines"
    },
    {
      "type": "table",
      "headers": ["Col1", "Col2"],
      "rows": [["data1", "data2"], ["data3", "data4"]]
    },
    {
      "type": "list",
      "ordered": false,
      "items": ["item 1", "item 2"]
    },
    {
      "type": "blockquote",
      "text": "quoted text"
    },
    {
      "type": "diagram",
      "title": "Diagram title describing what it shows",
      "description": "Brief explanation of the diagram purpose",
      "components": [
        { "name": "Component Name", "description": "What it does or represents" }
      ],
      "connections": [
        { "from": "Source", "to": "Target", "label": "relationship" }
      ]
    },
    {
      "type": "hr"
    }
  ]
}

CRITICAL RULES:
1. ASCII diagrams (content with box-drawing chars like ─│┌┐└┘├┤┬┴┼▼▲ or boxes made with +--+) MUST be converted to "diagram" type
2. For diagrams: Extract ALL components/boxes as named items with descriptions of what they do
3. For diagrams: Identify data flow connections between components (arrows, lines)
4. Preserve ALL code blocks exactly with language tags
5. Keep inline formatting (bold, italic, inline code, links) in paragraphs
6. Handle emojis - keep them in headings and text
7. Parse tables with all rows and columns
8. Create "hr" elements for horizontal rules (---)
9. Numbered lists should have "ordered": true

IMPORTANT: Respond with ONLY valid JSON, no markdown code blocks or explanations.`;

    let lastError = null;
    
    for (let attempt = 0; attempt <= RETRY_CONFIG.maxRetries; attempt++) {
        try {
            if (attempt > 0) {
                const suggestedDelay = extractRetryDelay(lastError);
                const delay = getRetryDelay(attempt - 1, suggestedDelay);
                console.log(`🔄 Retry attempt ${attempt}/${RETRY_CONFIG.maxRetries} after ${Math.round(delay / 1000)}s...`);
                await sleep(delay);
            }
            
            console.log('📤 Sending document to Gemini for parsing...');
            const result = await model.generateContent(prompt);
            const response = await result.response;
            let text = response.text().trim();
            
            // Clean up response
            text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
            
            const parsed = JSON.parse(text);
            console.log(`✅ Gemini parsed document with ${parsed.elements?.length || 0} elements`);
            
            return parsed;
        } catch (error) {
            lastError = error;
            const isRetryable = error.message?.includes('429') || 
                               error.message?.includes('503') || 
                               error.message?.includes('quota') ||
                               error.message?.includes('overloaded') ||
                               error.message?.includes('rate');
            
            console.error(`❌ Gemini error (attempt ${attempt + 1}): ${error.message?.substring(0, 150) || error}`);
            
            if (!isRetryable || attempt >= RETRY_CONFIG.maxRetries) {
                console.log('⚠️ Falling back to local parsing');
                return null;
            }
        }
    }
    
    return null;
}
