import MarkdownIt from 'markdown-it';
import hljs from 'highlight.js';
import { full as emoji } from 'markdown-it-emoji';
import anchor from 'markdown-it-anchor';

/**
 * Detect if content contains ASCII box-drawing characters
 */
function isAsciiDiagram(code) {
    const boxChars = /[─│┌┐└┘├┤┬┴┼═║╔╗╚╝╠╣╦╩╬▼▲◄►●○■□▪▫]/;
    return boxChars.test(code);
}

/**
 * Create and configure the markdown-it instance
 */
function createMarkdownParser() {
    const md = new MarkdownIt({
        html: true,
        linkify: true,
        typographer: true,
        breaks: false,
        highlight: function (str, lang) {
            // Check if it's an ASCII diagram (no lang or text/plaintext)
            if (!lang || lang === 'text' || lang === 'plaintext') {
                if (isAsciiDiagram(str)) {
                    return `<pre class="ascii-diagram"><code>${md.utils.escapeHtml(str)}</code></pre>`;
                }
            }
            
            // Syntax highlighting for code blocks
            if (lang && hljs.getLanguage(lang)) {
                try {
                    const highlighted = hljs.highlight(str, { 
                        language: lang, 
                        ignoreIllegals: true 
                    });
                    return `<pre class="hljs code-block" data-language="${lang}"><code>${highlighted.value}</code></pre>`;
                } catch (err) {
                    console.error('Highlight error:', err);
                }
            }
            
            // Fallback - no highlighting
            return `<pre class="code-block"><code>${md.utils.escapeHtml(str)}</code></pre>`;
        }
    });

    // Enable plugins
    md.use(emoji);
    md.use(anchor, {
        permalink: false,
        slugify: (s) => s.toLowerCase().replace(/[^\w]+/g, '-')
    });

    return md;
}

// Create singleton parser instance
const mdParser = createMarkdownParser();

/**
 * Parse markdown content to HTML
 * @param {string} markdown - The markdown content to parse
 * @returns {string} - The rendered HTML
 */
export function parseMarkdown(markdown) {
    if (!markdown || typeof markdown !== 'string') {
        return '';
    }
    return mdParser.render(markdown);
}

/**
 * Parse markdown to tokens (for DOCX conversion)
 * @param {string} markdown - The markdown content to parse
 * @returns {Array} - Array of tokens
 */
export function parseToTokens(markdown) {
    if (!markdown || typeof markdown !== 'string') {
        return [];
    }
    return mdParser.parse(markdown, {});
}

/**
 * Check if a code block is an ASCII diagram
 */
export { isAsciiDiagram };
