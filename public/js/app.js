// DOM Elements
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const clearBtn = document.getElementById('clearBtn');
const dropZone = document.getElementById('dropZone');
const loadingOverlay = document.getElementById('loadingOverlay');
const statusIndicator = document.getElementById('statusIndicator');
const statusText = document.getElementById('statusText');
const useAIToggle = document.getElementById('useAIToggle');

// State
let currentFilename = 'document.md';
let debounceTimer = null;
const STORAGE_KEY = 'md-to-docx-content';
const AI_TOGGLE_KEY = 'md-to-docx-ai-toggle';

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadFromCache();
    loadAIToggleState();
    setupEventListeners();
    // Trigger initial preview if there's content
    if (markdownInput.value.trim()) {
        updatePreview();
    }
});

// LocalStorage Functions
function saveToCache() {
    const markdown = markdownInput.value;
    if (markdown.trim()) {
        localStorage.setItem(STORAGE_KEY, markdown);
        localStorage.setItem(STORAGE_KEY + '_timestamp', Date.now());
        updateClearButtonState(true);
    }
}

function loadFromCache() {
    try {
        const cached = localStorage.getItem(STORAGE_KEY);
        if (cached) {
            markdownInput.value = cached;
            const timestamp = localStorage.getItem(STORAGE_KEY + '_timestamp');
            if (timestamp) {
                const date = new Date(parseInt(timestamp));
                console.log('✅ Loaded cached markdown from:', date.toLocaleString());
                setStatus('Loaded from cache', false);
                setTimeout(() => setStatus('Ready', false), 2000);
            }
            updateClearButtonState(true);
        }
    } catch (error) {
        console.error('Failed to load from cache:', error);
    }
}

function clearCache() {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(STORAGE_KEY + '_timestamp');
    updateClearButtonState(false);
}

function loadAIToggleState() {
    try {
        const saved = localStorage.getItem(AI_TOGGLE_KEY);
        if (saved !== null) {
            useAIToggle.checked = saved === 'true';
        }
    } catch (error) {
        console.error('Failed to load AI toggle state:', error);
    }
}

function saveAIToggleState() {
    try {
        localStorage.setItem(AI_TOGGLE_KEY, useAIToggle.checked);
    } catch (error) {
        console.error('Failed to save AI toggle state:', error);
    }
}

function updateClearButtonState(hasCached) {
    if (hasCached) {
        clearBtn.title = 'Clear editor and cached data';
    } else {
        clearBtn.title = 'Clear editor';
    }
}

// Event Listeners
function setupEventListeners() {
    // Live preview on input
    markdownInput.addEventListener('input', () => {
        debounce(updatePreview, 300);
        debounce(saveToCache, 1000); // Auto-save after 1 second of no typing
    });

    // File upload
    fileInput.addEventListener('change', handleFileSelect);

    // Download button
    downloadBtn.addEventListener('click', downloadDocx);

    // Clear button
    clearBtn.addEventListener('click', clearEditor);

    // AI Toggle - save state when changed
    useAIToggle.addEventListener('change', saveAIToggleState);

    // Drag and drop
    document.addEventListener('dragenter', showDropZone);
    document.addEventListener('dragleave', hideDropZone);
    document.addEventListener('dragover', handleDragOver);
    document.addEventListener('drop', handleDrop);

    // Paste handling
    markdownInput.addEventListener('paste', handlePaste);

    // Save on page unload
    window.addEventListener('beforeunload', () => {
        if (markdownInput.value.trim()) {
            saveToCache();
        }
    });
}

// Debounce function
function debounce(func, delay) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(func, delay);
}

// Update preview
async function updatePreview() {
    const markdown = markdownInput.value;
    
    if (!markdown.trim()) {
        preview.innerHTML = '<p class="placeholder-text">Your rendered Markdown will appear here...</p>';
        downloadBtn.disabled = true;
        setStatus('Ready', false);
        return;
    }

    setStatus('Updating...', 'loading');

    try {
        const response = await fetch('/api/preview', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ markdown })
        });

        if (!response.ok) {
            throw new Error('Failed to render preview');
        }

        const data = await response.json();
        preview.innerHTML = data.html;
        downloadBtn.disabled = false;
        setStatus('Ready', false);
    } catch (error) {
        console.error('Preview error:', error);
        setStatus('Error', 'error');
    }
}

// Handle file selection
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        loadFile(file);
    }
    // Reset input so same file can be selected again
    event.target.value = '';
}

// Load file content
async function loadFile(file) {
    setStatus('Loading...', 'loading');
    
    try {
        const formData = new FormData();
        formData.append('file', file);

        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Failed to upload file');
        }

        const data = await response.json();
        markdownInput.value = data.markdown;
        preview.innerHTML = data.html;
        currentFilename = data.filename;
        downloadBtn.disabled = false;
        setStatus('Ready', false);
        saveToCache(); // Save uploaded file to cache
    } catch (error) {
        console.error('Upload error:', error);
        setStatus('Error: ' + error.message, 'error');
        alert('Error loading file: ' + error.message);
    }
}

// Download DOCX
async function downloadDocx() {
    const markdown = markdownInput.value;
    
    if (!markdown.trim()) {
        alert('Please enter some Markdown content first.');
        return;
    }

    loadingOverlay.classList.remove('hidden');

    try {
        // Check if AI toggle exists and get its value
        const useAIToggle = document.getElementById('useAIToggle');
        const useAI = useAIToggle ? useAIToggle.checked : true;
        
        const response = await fetch('/api/convert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ 
                markdown, 
                filename: currentFilename,
                useAI: useAI
            })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Failed to convert to DOCX');
        }

        // Get filename from response header or use default
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'document.docx';
        if (contentDisposition) {
            const match = contentDisposition.match(/filename="(.+)"/);
            if (match) {
                filename = match[1];
            }
        }

        // Download the file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        setStatus('Downloaded!', false);
    } catch (error) {
        console.error('Download error:', error);
        setStatus('Error', 'error');
        alert('Error converting to DOCX: ' + error.message);
    } finally {
        loadingOverlay.classList.add('hidden');
    }
}

// Clear editor
function clearEditor() {
    if (markdownInput.value.trim() && !confirm('Are you sure you want to clear the editor and cached data?')) {
        return;
    }
    markdownInput.value = '';
    preview.innerHTML = '<p class="placeholder-text">Your rendered Markdown will appear here...</p>';
    downloadBtn.disabled = true;
    currentFilename = 'document.md';
    setStatus('Ready', false);
    clearCache(); // Clear localStorage
    console.log('Editor and cache cleared');
}

// Drag and drop handlers
function showDropZone(event) {
    event.preventDefault();
    if (event.dataTransfer.types.includes('Files')) {
        dropZone.classList.add('active');
    }
}

function hideDropZone(event) {
    event.preventDefault();
    // Only hide if we're leaving the window
    if (event.relatedTarget === null || !document.body.contains(event.relatedTarget)) {
        dropZone.classList.remove('active');
    }
}

function handleDragOver(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
}

function handleDrop(event) {
    event.preventDefault();
    dropZone.classList.remove('active');

    const files = event.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        const ext = file.name.split('.').pop().toLowerCase();
        
        if (['md', 'markdown', 'txt'].includes(ext)) {
            loadFile(file);
        } else {
            alert('Please drop a Markdown file (.md, .markdown, or .txt)');
        }
    }
}

// Handle paste
function handlePaste(event) {
    // Allow default paste behavior in textarea
    // Could enhance to handle pasted files if needed
}

// Set status
function setStatus(text, state) {
    statusText.textContent = text;
    statusIndicator.classList.remove('loading', 'error');
    
    if (state === 'loading') {
        statusIndicator.classList.add('loading');
    } else if (state === 'error') {
        statusIndicator.classList.add('error');
    }
}

// Resizable Panel Divider
const resizer = document.getElementById('resizer');
const editorPanel = document.querySelector('.editor-panel');
const previewPanel = document.querySelector('.preview-panel');

let isResizing = false;

resizer.addEventListener('mousedown', (e) => {
    isResizing = true;
    resizer.classList.add('resizing');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
});

document.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    
    const container = document.querySelector('.main-content');
    const containerRect = container.getBoundingClientRect();
    const containerWidth = containerRect.width;
    const offsetX = e.clientX - containerRect.left;
    
    // Calculate percentages (account for resizer width)
    const resizerWidth = 6;
    const minWidth = 200;
    
    // Clamp the position
    const clampedX = Math.max(minWidth, Math.min(offsetX, containerWidth - minWidth - resizerWidth));
    
    const editorWidth = clampedX;
    const previewWidth = containerWidth - clampedX - resizerWidth;
    
    editorPanel.style.flex = 'none';
    editorPanel.style.width = editorWidth + 'px';
    previewPanel.style.flex = 'none';
    previewPanel.style.width = previewWidth + 'px';
});

document.addEventListener('mouseup', () => {
    if (isResizing) {
        isResizing = false;
        resizer.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    }
});
