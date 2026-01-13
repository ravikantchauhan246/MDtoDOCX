import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    ShadingType,
    convertInchesToTwip,
    UnderlineType
} from 'docx';
import { parseToTokens, isAsciiDiagram } from '../parser/markdownParser.js';
import { parseDocumentWithAI, convertDiagramWithAI, isGeminiAvailable } from '../services/geminiService.js';

// Style configuration
const STYLES = {
    fonts: {
        heading: 'Calibri',
        body: 'Calibri',
        code: 'Consolas'
    },
    sizes: {
        h1: 28,
        h2: 24,
        h3: 20,
        h4: 18,
        h5: 16,
        h6: 14,
        body: 11,
        code: 10,
        table: 10
    },
    colors: {
        heading: '1a1a1a',
        body: '333333',
        code: '24292e',
        codeBackground: 'f6f8fa',
        link: '0366d6',
        tableHeader: 'f6f8fa',
        tableBorder: 'd0d7de',
        diagramBg: 'e8f4fd',
        diagramBorder: '0366d6',
        componentBg: 'ffffff'
    }
};

/**
 * Convert markdown to DOCX buffer
 * Tries AI parsing first, falls back to token-based parsing
 */
export async function convertToDocx(markdown, useAI = true) {
    let docElements = [];
    let usedAI = false;
    
    // Try AI parsing first if enabled
    if (useAI && isGeminiAvailable()) {
        try {
            const aiParsed = await parseDocumentWithAI(markdown);
            if (aiParsed && aiParsed.elements && aiParsed.elements.length > 0) {
                console.log(`🤖 Using AI-parsed structure (${aiParsed.elements.length} elements)`);
                docElements = await processAIParsedElements(aiParsed);
                usedAI = true;
            }
        } catch (error) {
            console.error('AI parsing failed, using fallback:', error.message);
        }
    }
    
    // Fallback to token-based parsing
    if (!usedAI) {
        console.log('📝 Using local token-based parsing');
        const tokens = parseToTokens(markdown);
        await processTokens(tokens, docElements);
    }
    
    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: convertInchesToTwip(1),
                        right: convertInchesToTwip(1),
                        bottom: convertInchesToTwip(1),
                        left: convertInchesToTwip(1)
                    }
                }
            },
            children: docElements
        }]
    });
    
    return await Packer.toBuffer(doc);
}

/**
 * Process AI-parsed document structure into DOCX elements
 */
async function processAIParsedElements(parsed) {
    const elements = [];
    
    for (const elem of parsed.elements) {
        switch (elem.type) {
            case 'heading':
                elements.push(createHeading(elem.text, elem.level || 1));
                break;
                
            case 'paragraph':
                elements.push(createAIParagraph(elem.content));
                break;
                
            case 'code':
                elements.push(...createCodeBlock(elem.content, elem.language || ''));
                break;
                
            case 'table':
                elements.push(createAITable(elem.headers, elem.rows));
                break;
                
            case 'list':
                elements.push(...createAIList(elem.items, elem.ordered));
                break;
                
            case 'blockquote':
                elements.push(createAIBlockquote(elem.text));
                break;
                
            case 'diagram':
                elements.push(...createAIDiagram(elem));
                break;
                
            case 'hr':
                elements.push(createHorizontalRule());
                break;
                
            default:
                // Handle unknown types as paragraphs
                if (elem.text) {
                    elements.push(new Paragraph({
                        spacing: { after: 160 },
                        children: [new TextRun({ text: elem.text })]
                    }));
                }
        }
    }
    
    return elements;
}

/**
 * Create paragraph from AI-parsed content with inline formatting
 */
function createAIParagraph(content) {
    if (!content || !Array.isArray(content)) {
        return new Paragraph({ spacing: { after: 160 }, children: [] });
    }
    
    const runs = content.map(item => {
        if (item.code) {
            return new TextRun({
                text: item.text,
                font: STYLES.fonts.code,
                size: STYLES.sizes.code * 2,
                shading: {
                    type: ShadingType.SOLID,
                    color: STYLES.colors.codeBackground
                }
            });
        }
        
        if (item.link) {
            return new TextRun({
                text: item.text,
                color: STYLES.colors.link,
                underline: { type: UnderlineType.SINGLE }
            });
        }
        
        return new TextRun({
            text: item.text,
            bold: item.bold || false,
            italics: item.italic || false,
            strike: item.strikethrough || false,
            font: STYLES.fonts.body,
            size: STYLES.sizes.body * 2,
            color: STYLES.colors.body
        });
    });
    
    return new Paragraph({
        spacing: { after: 160 },
        children: runs
    });
}

/**
 * Create table from AI-parsed headers and rows
 */
function createAITable(headers, rows) {
    const tableRows = [];
    
    // Header row
    if (headers && headers.length > 0) {
        tableRows.push(new TableRow({
            children: headers.map(h => new TableCell({
                children: [new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [new TextRun({
                        text: h,
                        bold: true,
                        font: STYLES.fonts.body,
                        size: STYLES.sizes.table * 2,
                        color: '24292f'
                    })]
                })],
                shading: { type: ShadingType.SOLID, color: STYLES.colors.tableHeader },
                margins: { top: 100, bottom: 100, left: 150, right: 150 },
                width: { size: 0, type: WidthType.AUTO }
            }))
        }));
    }
    
    // Data rows
    if (rows && rows.length > 0) {
        for (const row of rows) {
            tableRows.push(new TableRow({
                children: row.map(cell => new TableCell({
                    children: [new Paragraph({
                        spacing: { before: 40, after: 40 },
                        children: [new TextRun({
                            text: cell || '',
                            font: STYLES.fonts.body,
                            size: STYLES.sizes.table * 2,
                            color: STYLES.colors.body
                        })]
                    })],
                    margins: { top: 100, bottom: 100, left: 150, right: 150 },
                    width: { size: 0, type: WidthType.AUTO }
                }))
            }));
        }
    }
    
    return new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: 'autofit',
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            left: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            right: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder }
        }
    });
}

/**
 * Create list from AI-parsed items
 */
function createAIList(items, ordered = false) {
    if (!items || !Array.isArray(items)) return [];
    
    return items.map((item, i) => {
        const bullet = ordered ? `${i + 1}. ` : '• ';
        return new Paragraph({
            spacing: { after: 80 },
            indent: { left: convertInchesToTwip(0.25) },
            children: [
                new TextRun({
                    text: bullet + item,
                    font: STYLES.fonts.body,
                    size: STYLES.sizes.body * 2
                })
            ]
        });
    });
}

/**
 * Create blockquote from AI-parsed text
 */
function createAIBlockquote(text) {
    return new Paragraph({
        spacing: { after: 80 },
        indent: { left: convertInchesToTwip(0.5) },
        border: {
            left: {
                style: BorderStyle.SINGLE,
                size: 24,
                color: 'd0d7de'
            }
        },
        children: [
            new TextRun({
                text: '  ' + text,
                font: STYLES.fonts.body,
                size: STYLES.sizes.body * 2,
                color: '6a737d',
                italics: true
            })
        ]
    });
}

/**
 * Create diagram block from AI-parsed diagram data
 */
function createAIDiagram(diagram) {
    const elements = [];
    
    // Add diagram title
    if (diagram.title) {
        elements.push(new Paragraph({
            spacing: { before: 240, after: 120 },
            shading: {
                type: ShadingType.SOLID,
                color: STYLES.colors.diagramBorder
            },
            children: [
                new TextRun({
                    text: ` 📊 ${diagram.title}`,
                    bold: true,
                    size: 24 * 2,
                    font: STYLES.fonts.heading,
                    color: 'ffffff'
                })
            ]
        }));
    }
    
    // Add description
    if (diagram.description) {
        elements.push(new Paragraph({
            spacing: { after: 120 },
            shading: {
                type: ShadingType.SOLID,
                color: STYLES.colors.diagramBg
            },
            children: [
                new TextRun({
                    text: ` ${diagram.description}`,
                    italics: true,
                    size: STYLES.sizes.body * 2,
                    font: STYLES.fonts.body,
                    color: '555555'
                })
            ]
        }));
    }
    
    // Create components table if available
    if (diagram.components && diagram.components.length > 0) {
        elements.push(createComponentsTable(diagram.components));
    }
    
    // Add connections/flow description
    if (diagram.connections && diagram.connections.length > 0) {
        elements.push(new Paragraph({
            spacing: { before: 160, after: 80 },
            children: [
                new TextRun({
                    text: '🔗 Data Flow:',
                    bold: true,
                    size: STYLES.sizes.body * 2,
                    font: STYLES.fonts.heading
                })
            ]
        }));
        
        for (const conn of diagram.connections) {
            const label = conn.label ? ` (${conn.label})` : '';
            elements.push(new Paragraph({
                spacing: { after: 60 },
                indent: { left: convertInchesToTwip(0.25) },
                children: [
                    new TextRun({
                        text: `→ ${conn.from}  ➜  ${conn.to}${label}`,
                        size: STYLES.sizes.body * 2,
                        font: STYLES.fonts.body,
                        color: STYLES.colors.body
                    })
                ]
            }));
        }
    }
    
    // Add spacing after diagram
    elements.push(new Paragraph({ spacing: { after: 200 } }));
    
    return elements;
}

/**
 * Process markdown tokens into DOCX elements
 */
async function processTokens(tokens, elements, context = {}) {
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        
        switch (token.type) {
            case 'heading_open':
                const headingContent = tokens[i + 1];
                const level = parseInt(token.tag.substring(1));
                elements.push(createHeading(headingContent.content, level));
                i += 2;
                break;
                
            case 'paragraph_open':
                const paragraphContent = tokens[i + 1];
                if (paragraphContent && paragraphContent.type === 'inline') {
                    elements.push(createParagraph(paragraphContent.children));
                }
                i += 2;
                break;
                
            case 'fence':
                // Check if it's an ASCII diagram
                if (isAsciiDiagram(token.content)) {
                    const diagramElements = await createDiagramBlock(token.content);
                    elements.push(...diagramElements);
                } else {
                    elements.push(...createCodeBlock(token.content, token.info));
                }
                break;
                
            case 'code_block':
                if (isAsciiDiagram(token.content)) {
                    const diagramElements = await createDiagramBlock(token.content);
                    elements.push(...diagramElements);
                } else {
                    elements.push(...createCodeBlock(token.content, ''));
                }
                break;
                
            case 'table_open':
                const tableEnd = findClosingToken(tokens, i, 'table_close');
                const tableTokens = tokens.slice(i, tableEnd + 1);
                elements.push(createTable(tableTokens));
                i = tableEnd;
                break;
                
            case 'bullet_list_open':
            case 'ordered_list_open':
                const listEnd = findClosingToken(tokens, i, token.type.replace('_open', '_close'));
                const listTokens = tokens.slice(i, listEnd + 1);
                const isOrdered = token.type === 'ordered_list_open';
                elements.push(...createList(listTokens, isOrdered));
                i = listEnd;
                break;
                
            case 'hr':
                elements.push(createHorizontalRule());
                break;
                
            case 'blockquote_open':
                const bqEnd = findClosingToken(tokens, i, 'blockquote_close');
                const bqTokens = tokens.slice(i + 1, bqEnd);
                elements.push(...createBlockquote(bqTokens));
                i = bqEnd;
                break;
        }
    }
}

/**
 * Find the closing token for a given opening token
 */
function findClosingToken(tokens, startIndex, closeType) {
    let depth = 1;
    const openType = closeType.replace('_close', '_open');
    
    for (let i = startIndex + 1; i < tokens.length; i++) {
        if (tokens[i].type === openType) depth++;
        if (tokens[i].type === closeType) depth--;
        if (depth === 0) return i;
    }
    return tokens.length - 1;
}

/**
 * Create a heading paragraph
 */
function createHeading(text, level) {
    const headingLevels = {
        1: HeadingLevel.HEADING_1,
        2: HeadingLevel.HEADING_2,
        3: HeadingLevel.HEADING_3,
        4: HeadingLevel.HEADING_4,
        5: HeadingLevel.HEADING_5,
        6: HeadingLevel.HEADING_6
    };
    
    const sizes = {
        1: STYLES.sizes.h1,
        2: STYLES.sizes.h2,
        3: STYLES.sizes.h3,
        4: STYLES.sizes.h4,
        5: STYLES.sizes.h5,
        6: STYLES.sizes.h6
    };
    
    return new Paragraph({
        heading: headingLevels[level] || HeadingLevel.HEADING_1,
        spacing: { before: 240, after: 120 },
        children: [
            new TextRun({
                text: cleanText(text),
                bold: true,
                size: sizes[level] * 2,
                font: STYLES.fonts.heading,
                color: STYLES.colors.heading
            })
        ]
    });
}

/**
 * Create a paragraph with inline formatting
 */
function createParagraph(inlineTokens) {
    const children = processInlineTokens(inlineTokens || []);
    
    return new Paragraph({
        spacing: { after: 160 },
        children: children.length > 0 ? children : [new TextRun({ text: '' })]
    });
}

/**
 * Process inline tokens (bold, italic, links, code, etc.)
 */
function processInlineTokens(tokens) {
    const runs = [];
    let formatting = { bold: false, italic: false, code: false, strikethrough: false };
    
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        
        switch (token.type) {
            case 'text':
                runs.push(createTextRun(token.content, formatting));
                break;
                
            case 'emoji':
                runs.push(createTextRun(token.content, formatting));
                break;
                
            case 'code_inline':
                runs.push(new TextRun({
                    text: token.content,
                    font: STYLES.fonts.code,
                    size: STYLES.sizes.code * 2,
                    shading: {
                        type: ShadingType.SOLID,
                        color: STYLES.colors.codeBackground
                    }
                }));
                break;
                
            case 'strong_open':
                formatting.bold = true;
                break;
            case 'strong_close':
                formatting.bold = false;
                break;
                
            case 'em_open':
                formatting.italic = true;
                break;
            case 'em_close':
                formatting.italic = false;
                break;
                
            case 's_open':
                formatting.strikethrough = true;
                break;
            case 's_close':
                formatting.strikethrough = false;
                break;
                
            case 'link_open':
                // Find the link text
                const linkText = tokens[i + 1];
                if (linkText && linkText.type === 'text') {
                    runs.push(new TextRun({
                        text: linkText.content,
                        color: STYLES.colors.link,
                        underline: { type: UnderlineType.SINGLE }
                    }));
                    i++; // Skip the text token
                }
                break;
                
            case 'softbreak':
            case 'hardbreak':
                runs.push(new TextRun({ text: '', break: 1 }));
                break;
        }
    }
    
    return runs;
}

/**
 * Create a text run with formatting
 */
function createTextRun(text, formatting = {}) {
    return new TextRun({
        text: text,
        bold: formatting.bold,
        italics: formatting.italic,
        strike: formatting.strikethrough,
        font: STYLES.fonts.body,
        size: STYLES.sizes.body * 2,
        color: STYLES.colors.body
    });
}

/**
 * Create a diagram block - converts ASCII diagram to formatted DOCX components
 */
async function createDiagramBlock(asciiContent) {
    const elements = [];
    
    try {
        // Use Gemini to analyze and convert the diagram
        const analysis = await convertDiagramWithAI(asciiContent);
        
        // Add diagram title if available
        if (analysis.title) {
            elements.push(new Paragraph({
                spacing: { before: 240, after: 120 },
                shading: {
                    type: ShadingType.SOLID,
                    color: STYLES.colors.diagramBorder
                },
                children: [
                    new TextRun({
                        text: ` 📊 ${analysis.title}`,
                        bold: true,
                        size: 24 * 2,
                        font: STYLES.fonts.heading,
                        color: 'ffffff'
                    })
                ]
            }));
        }
        
        // Add description
        if (analysis.description) {
            elements.push(new Paragraph({
                spacing: { after: 120 },
                shading: {
                    type: ShadingType.SOLID,
                    color: STYLES.colors.diagramBg
                },
                children: [
                    new TextRun({
                        text: ` ${analysis.description}`,
                        italics: true,
                        size: STYLES.sizes.body * 2,
                        font: STYLES.fonts.body,
                        color: '555555'
                    })
                ]
            }));
        }
        
        // Create a table for components if available
        if (analysis.components && analysis.components.length > 0) {
            elements.push(createComponentsTable(analysis.components));
        }
        
        // Add connections/flow description
        if (analysis.connections && analysis.connections.length > 0) {
            elements.push(new Paragraph({
                spacing: { before: 160, after: 80 },
                children: [
                    new TextRun({
                        text: '🔗 Data Flow:',
                        bold: true,
                        size: STYLES.sizes.body * 2,
                        font: STYLES.fonts.heading
                    })
                ]
            }));
            
            for (const conn of analysis.connections) {
                const label = conn.label ? ` (${conn.label})` : '';
                elements.push(new Paragraph({
                    spacing: { after: 60 },
                    indent: { left: convertInchesToTwip(0.25) },
                    children: [
                        new TextRun({
                            text: `→ ${conn.from}  ➜  ${conn.to}${label}`,
                            size: STYLES.sizes.body * 2,
                            font: STYLES.fonts.body,
                            color: STYLES.colors.body
                        })
                    ]
                }));
            }
        }
        
        // Add summary if no components
        if ((!analysis.components || analysis.components.length === 0) && analysis.summary) {
            const summaryLines = analysis.summary.split('\n').filter(l => l.trim());
            for (const line of summaryLines) {
                elements.push(new Paragraph({
                    spacing: { after: 60 },
                    shading: {
                        type: ShadingType.SOLID,
                        color: STYLES.colors.diagramBg
                    },
                    indent: { left: convertInchesToTwip(0.25) },
                    children: [
                        new TextRun({
                            text: line.startsWith('•') ? ` ${line}` : ` • ${line}`,
                            size: STYLES.sizes.body * 2,
                            font: STYLES.fonts.body
                        })
                    ]
                }));
            }
        }
        
        // Add spacing after diagram
        elements.push(new Paragraph({ spacing: { after: 200 } }));
        
    } catch (error) {
        console.error('Diagram conversion error:', error);
        // Fallback: render as monospace text
        elements.push(...createCodeBlock(asciiContent, 'text'));
    }
    
    return elements;
}

/**
 * Create a table showing diagram components
 */
function createComponentsTable(components) {
    const rows = [];
    
    // Header row
    rows.push(new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({ text: 'Component', bold: true, size: STYLES.sizes.body * 2, color: 'ffffff' })]
                })],
                shading: { type: ShadingType.SOLID, color: STYLES.colors.diagramBorder },
                margins: { top: 80, bottom: 80, left: 120, right: 120 }
            }),
            new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({ text: 'Description', bold: true, size: STYLES.sizes.body * 2, color: 'ffffff' })]
                })],
                shading: { type: ShadingType.SOLID, color: STYLES.colors.diagramBorder },
                margins: { top: 80, bottom: 80, left: 120, right: 120 }
            })
        ]
    }));
    
    // Component rows
    for (const comp of components) {
        rows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({
                        children: [new TextRun({ 
                            text: comp.name || '', 
                            bold: true,
                            size: STYLES.sizes.body * 2,
                            color: STYLES.colors.diagramBorder
                        })]
                    })],
                    shading: { type: ShadingType.SOLID, color: STYLES.colors.diagramBg },
                    margins: { top: 60, bottom: 60, left: 120, right: 120 }
                }),
                new TableCell({
                    children: [new Paragraph({
                        children: [new TextRun({ 
                            text: comp.description || '-', 
                            size: STYLES.sizes.body * 2 
                        })]
                    })],
                    margins: { top: 60, bottom: 60, left: 120, right: 120 }
                })
            ]
        }));
    }
    
    return new Table({
        rows: rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.diagramBorder },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.diagramBorder },
            left: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.diagramBorder },
            right: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.diagramBorder },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder }
        }
    });
}

/**
 * Create code block paragraphs
 */
function createCodeBlock(code, language) {
    const lines = code.split('\n');
    const paragraphs = [];
    
    // Add language label if present
    if (language) {
        paragraphs.push(new Paragraph({
            spacing: { before: 120, after: 0 },
            shading: {
                type: ShadingType.SOLID,
                color: 'd0d7de'
            },
            children: [
                new TextRun({
                    text: `  ${language.toUpperCase()}`,
                    font: STYLES.fonts.code,
                    size: 16 * 2,
                    bold: true,
                    color: '57606a'
                })
            ]
        }));
    }
    
    for (const line of lines) {
        paragraphs.push(new Paragraph({
            spacing: { 
                after: 0, 
                before: 0,
                line: 260
            },
            shading: {
                type: ShadingType.SOLID,
                color: STYLES.colors.codeBackground
            },
            children: [
                new TextRun({
                    text: '  ' + (line || ' '),
                    font: STYLES.fonts.code,
                    size: STYLES.sizes.code * 2,
                    color: STYLES.colors.code
                })
            ]
        }));
    }
    
    // Add spacing after code block
    paragraphs.push(new Paragraph({ spacing: { after: 160 } }));
    
    return paragraphs;
}

/**
 * Create a table from tokens
 */
function createTable(tokens) {
    const rows = [];
    let currentRow = [];
    let isHeader = false;
    let inCell = false;
    let cellContent = [];
    
    for (const token of tokens) {
        switch (token.type) {
            case 'thead_open':
                isHeader = true;
                break;
            case 'thead_close':
                isHeader = false;
                break;
            case 'tr_open':
                currentRow = [];
                break;
            case 'tr_close':
                if (currentRow.length > 0) {
                    rows.push({ cells: currentRow, isHeader });
                }
                break;
            case 'th_open':
            case 'td_open':
                inCell = true;
                cellContent = [];
                break;
            case 'th_close':
            case 'td_close':
                inCell = false;
                currentRow.push({
                    content: cellContent.join(''),
                    isHeader: token.type === 'th_close'
                });
                break;
            case 'inline':
                if (inCell) {
                    cellContent.push(extractTextFromInline(token.children));
                }
                break;
        }
    }
    
    if (rows.length === 0) {
        return new Paragraph({ text: '' });
    }
    
    const columnCount = Math.max(...rows.map(r => r.cells.length));
    
    const tableRows = rows.map((row, rowIndex) => {
        const cells = row.cells.map((cell, cellIndex) => {
            return new TableCell({
                children: [
                    new Paragraph({
                        spacing: { before: 40, after: 40 },
                        children: [
                            new TextRun({
                                text: cell.content,
                                bold: cell.isHeader,
                                font: STYLES.fonts.body,
                                size: STYLES.sizes.table * 2,
                                color: cell.isHeader ? '24292f' : STYLES.colors.body
                            })
                        ]
                    })
                ],
                shading: cell.isHeader ? {
                    type: ShadingType.SOLID,
                    color: STYLES.colors.tableHeader
                } : undefined,
                margins: {
                    top: 100,
                    bottom: 100,
                    left: 150,
                    right: 150
                },
                width: { size: 0, type: WidthType.AUTO }
            });
        });
        
        // Pad row if needed
        while (cells.length < columnCount) {
            cells.push(new TableCell({
                children: [new Paragraph({ text: '' })]
            }));
        }
        
        return new TableRow({ children: cells });
    });
    
    return new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: 'autofit',
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            left: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            right: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: STYLES.colors.tableBorder }
        }
    });
}

/**
 * Extract plain text from inline tokens
 */
function extractTextFromInline(tokens) {
    if (!tokens) return '';
    return tokens
        .filter(t => t.type === 'text' || t.type === 'code_inline' || t.type === 'emoji')
        .map(t => t.content)
        .join('');
}

/**
 * Create list items
 */
function createList(tokens, isOrdered) {
    const items = [];
    let itemNumber = 1;
    let currentContent = '';
    
    for (const token of tokens) {
        if (token.type === 'list_item_open') {
            currentContent = '';
        } else if (token.type === 'inline') {
            currentContent = extractTextFromInline(token.children);
        } else if (token.type === 'list_item_close') {
            const bullet = isOrdered ? `${itemNumber}. ` : '• ';
            items.push(new Paragraph({
                spacing: { after: 80 },
                indent: { left: convertInchesToTwip(0.25) },
                children: [
                    new TextRun({
                        text: bullet + currentContent,
                        font: STYLES.fonts.body,
                        size: STYLES.sizes.body * 2
                    })
                ]
            }));
            itemNumber++;
        }
    }
    
    return items;
}

/**
 * Create horizontal rule
 */
function createHorizontalRule() {
    return new Paragraph({
        spacing: { before: 200, after: 200 },
        border: {
            bottom: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: 'd0d7de'
            }
        },
        children: [new TextRun({ text: '' })]
    });
}

/**
 * Create blockquote paragraphs
 */
function createBlockquote(tokens) {
    const paragraphs = [];
    
    for (const token of tokens) {
        if (token.type === 'paragraph_open') continue;
        if (token.type === 'paragraph_close') continue;
        
        if (token.type === 'inline') {
            paragraphs.push(new Paragraph({
                spacing: { after: 80 },
                indent: { left: convertInchesToTwip(0.5) },
                border: {
                    left: {
                        style: BorderStyle.SINGLE,
                        size: 24,
                        color: 'd0d7de'
                    }
                },
                children: [
                    new TextRun({
                        text: '  ' + extractTextFromInline(token.children),
                        font: STYLES.fonts.body,
                        size: STYLES.sizes.body * 2,
                        color: '6a737d',
                        italics: true
                    })
                ]
            }));
        }
    }
    
    return paragraphs;
}

/**
 * Clean text - remove markdown artifacts
 */
function cleanText(text) {
    if (!text) return '';
    return text
        .replace(/\*\*/g, '')
        .replace(/\*/g, '')
        .replace(/`/g, '')
        .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
        .trim();
}
