Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Initialize tab functionality
        initializeTabs();
        
        // Add form event listener
        document.getElementById("titlePageForm").addEventListener("submit", insertTitlePage);
        
        // Add table generator event listeners
        initializeTableGenerator();
        
        // Load saved preferences
        loadUserPreferences();
    }
});

// Initialize user preferences with defaults
window.userPreferences = {
    defaultTextColor: '#000000',
    defaultTextAlignment: 'left',
    defaultVerticalAlignment: 'center',
    defaultCellBgColor: '#ffffff',
    defaultHeaderBgColor: '#f8f9fa'
};

// Load user preferences from Office settings and localStorage
function loadUserPreferences() {
    try {
        console.log('Loading preferences...');
        let preferences = null;
        
        // Try Office.context.document.settings first
        if (Office.context && Office.context.document && Office.context.document.settings) {
            try {
                preferences = Office.context.document.settings.get('wordAddinPreferences');
                console.log('Raw preferences from Office settings:', preferences);
            } catch (officeError) {
                console.log('Office settings not available, trying localStorage');
            }
        }
        
        // Fallback to localStorage
        if (!preferences) {
            preferences = localStorage.getItem('wordAddinPreferences');
            console.log('Raw preferences from localStorage:', preferences);
        }
        
        if (preferences) {
            const prefs = JSON.parse(preferences);
            console.log('Parsed preferences:', prefs);
            
            // Merge with defaults
            window.userPreferences = { ...window.userPreferences, ...prefs };
            console.log('Updated window.userPreferences:', window.userPreferences);
            
            // Apply to UI
            applyPreferencesToUI();
        } else {
            console.log('No preferences found in any storage');
        }
        updateFormatStatus();
    } catch (error) {
        console.log('Could not load preferences:', error);
    }
}

// Apply preferences to UI elements
function applyPreferencesToUI() {
    // Use a small delay to ensure DOM is ready
    setTimeout(() => {
        const textColor = document.getElementById('textColor');
        const textAlignment = document.getElementById('textAlignment');
        const verticalAlignment = document.getElementById('verticalAlignment');
        const cellBgColor = document.getElementById('cellBgColor');
        const headerBgColor = document.getElementById('headerBgColor');
        
        console.log('Applying preferences:', window.userPreferences);
        console.log('Found elements:', {
            textColor: !!textColor,
            textAlignment: !!textAlignment,
            verticalAlignment: !!verticalAlignment,
            cellBgColor: !!cellBgColor,
            headerBgColor: !!headerBgColor
        });
        
        if (textColor) {
            textColor.value = window.userPreferences.defaultTextColor;
        }
        if (textAlignment) {
            setAlignmentButtonActive('textAlignment', window.userPreferences.defaultTextAlignment);
        }
        if (verticalAlignment) {
            setAlignmentButtonActive('verticalAlignment', window.userPreferences.defaultVerticalAlignment);
        }
        if (cellBgColor) {
            cellBgColor.value = window.userPreferences.defaultCellBgColor;
        }
        if (headerBgColor) {
            headerBgColor.value = window.userPreferences.defaultHeaderBgColor;
        }
        
        // Apply formatting to any existing table preview
        applyFormattingToPreview();
        
        // Initialize alignment buttons with default states
        setAlignmentButtonActive('textAlignment', window.userPreferences.defaultTextAlignment);
        setAlignmentButtonActive('verticalAlignment', window.userPreferences.defaultVerticalAlignment);
    }, 100);
}

// Save user preferences using Office settings API and localStorage as fallback
async function saveUserPreferences(preferences) {
    try {
        console.log('Saving preferences:', preferences);
        
        // Get existing preferences
        let savedPrefs = {};
        
        // Try Office.context.document.settings first (if available)
        if (Office.context && Office.context.document && Office.context.document.settings) {
            try {
                const existingOfficePrefs = Office.context.document.settings.get('wordAddinPreferences');
                if (existingOfficePrefs) {
                    savedPrefs = JSON.parse(existingOfficePrefs);
                }
            } catch (officeError) {
                console.log('Office settings not available, falling back to localStorage');
            }
        }
        
        // Fallback to localStorage
        if (Object.keys(savedPrefs).length === 0) {
            const existingPreferences = localStorage.getItem('wordAddinPreferences');
            console.log('Existing preferences in localStorage:', existingPreferences);
            
            if (existingPreferences) {
                savedPrefs = JSON.parse(existingPreferences);
            }
        }
        
        console.log('Parsed existing preferences:', savedPrefs);
        
        // Merge new preferences with existing ones
        const mergedPreferences = { ...savedPrefs, ...preferences };
        console.log('Merged preferences:', mergedPreferences);
        
        // Save to Office settings if available
        if (Office.context && Office.context.document && Office.context.document.settings) {
            try {
                Office.context.document.settings.set('wordAddinPreferences', JSON.stringify(mergedPreferences));
                await Office.context.document.settings.saveAsync();
                console.log('Saved to Office settings');
            } catch (officeError) {
                console.log('Could not save to Office settings:', officeError);
            }
        }
        
        // Also save to localStorage as backup
        localStorage.setItem('wordAddinPreferences', JSON.stringify(mergedPreferences));
        console.log('Saved to localStorage. Verifying...');
        
        // Verify it was saved
        const verification = localStorage.getItem('wordAddinPreferences');
        console.log('Verification - what was actually saved:', verification);
        
        window.userPreferences = { ...window.userPreferences, ...mergedPreferences };
        console.log('Updated window.userPreferences:', window.userPreferences);
        
        updateFormatStatus();
        showMessage('Préférences sauvegardées!', 'success');
    } catch (error) {
        console.error('Could not save preferences:', error);
        showMessage('Erreur lors de la sauvegarde', 'error');
    }
}

// Update format status indicators
function updateFormatStatus() {
    const statusElements = [
        { id: 'textColorStatus', key: 'defaultTextColor', defaultValue: '#000000' },
        { id: 'textAlignmentStatus', key: 'defaultTextAlignment', defaultValue: 'left' },
        { id: 'verticalAlignmentStatus', key: 'defaultVerticalAlignment', defaultValue: 'center' },
        { id: 'cellBgColorStatus', key: 'defaultCellBgColor', defaultValue: '#ffffff' },
        { id: 'headerBgColorStatus', key: 'defaultHeaderBgColor', defaultValue: '#f8f9fa' }
    ];
    
    const saved = localStorage.getItem('wordAddinPreferences');
    if (saved) {
        const prefs = JSON.parse(saved);
        statusElements.forEach(({ id, key, defaultValue }) => {
            const element = document.getElementById(id);
            if (element) {
                const isCustom = prefs[key] !== undefined && prefs[key] !== defaultValue;
                element.textContent = isCustom ? 'Préférence sauvegardée' : 'Valeur par défaut';
                element.className = 'format-status ' + (isCustom ? 'saved' : 'default');
            }
        });
    } else {
        // No saved preferences, all are default
        statusElements.forEach(({ id }) => {
            const element = document.getElementById(id);
            if (element) {
                element.textContent = 'Valeur par défaut';
                element.className = 'format-status default';
            }
        });
    }
}

// Tab switching functionality
function initializeTabs() {
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');
    
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.getAttribute('data-tab');
            
            // Remove active class from all buttons and contents
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));
            
            // Add active class to clicked button and corresponding content
            button.classList.add('active');
            document.getElementById(targetTab).classList.add('active');
        });
    });
}

// Table generator functionality
function initializeTableGenerator() {
    const generateBtn = document.getElementById('generateTablePreview');
    const insertBtn = document.getElementById('insertTable');
    const resetBtn = document.getElementById('resetTable');
    
    if (generateBtn) generateBtn.addEventListener('click', generateTablePreview);
    if (insertBtn) insertBtn.addEventListener('click', insertTableToWord);
    if (resetBtn) resetBtn.addEventListener('click', resetTableGenerator);
    
    // Initialize formatting buttons
    initializeFormattingButtons();
    
    // Initialize format controls functionality
    initializeFormatControls();
}

// Set active alignment button
function setAlignmentButtonActive(containerId, alignment) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    // Update data attribute
    container.setAttribute('data-alignment', alignment);
    
    // Update button states
    const buttons = container.querySelectorAll('.alignment-btn');
    buttons.forEach(btn => {
        if (btn.getAttribute('data-align') === alignment) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

// Get current alignment value
function getAlignmentValue(containerId) {
    const container = document.getElementById(containerId);
    return container ? container.getAttribute('data-alignment') : 'left';
}

// Initialize format control event handlers
function initializeFormatControls() {
    // Save button handlers
    const saveTextColorBtn = document.getElementById('saveTextColor');
    const saveTextAlignmentBtn = document.getElementById('saveTextAlignment');
    const saveVerticalAlignmentBtn = document.getElementById('saveVerticalAlignment');
    const saveCellBgColorBtn = document.getElementById('saveCellBgColor');
    const saveHeaderBgColorBtn = document.getElementById('saveHeaderBgColor');
    
    if (saveTextColorBtn) {
        saveTextColorBtn.addEventListener('click', () => {
            const color = document.getElementById('textColor').value;
            saveUserPreferences({ defaultTextColor: color });
        });
    }
    
    if (saveTextAlignmentBtn) {
        saveTextAlignmentBtn.addEventListener('click', () => {
            const alignment = getAlignmentValue('textAlignment');
            saveUserPreferences({ defaultTextAlignment: alignment });
        });
    }
    
    if (saveVerticalAlignmentBtn) {
        saveVerticalAlignmentBtn.addEventListener('click', () => {
            const alignment = getAlignmentValue('verticalAlignment');
            saveUserPreferences({ defaultVerticalAlignment: alignment });
        });
    }
    
    if (saveCellBgColorBtn) {
        saveCellBgColorBtn.addEventListener('click', () => {
            const color = document.getElementById('cellBgColor').value;
            saveUserPreferences({ defaultCellBgColor: color });
        });
    }
    
    if (saveHeaderBgColorBtn) {
        saveHeaderBgColorBtn.addEventListener('click', () => {
            const color = document.getElementById('headerBgColor').value;
            saveUserPreferences({ defaultHeaderBgColor: color });
        });
    }
    
    // Reset button handlers
    const resetTextColorBtn = document.getElementById('resetTextColor');
    const resetCellBgColorBtn = document.getElementById('resetCellBgColor');
    const resetHeaderBgColorBtn = document.getElementById('resetHeaderBgColor');
    
    if (resetTextColorBtn) {
        resetTextColorBtn.addEventListener('click', () => {
            document.getElementById('textColor').value = '#000000';
            applyFormattingToPreview();
        });
    }
    
    if (resetCellBgColorBtn) {
        resetCellBgColorBtn.addEventListener('click', () => {
            document.getElementById('cellBgColor').value = '#ffffff';
            applyFormattingToPreview();
        });
    }
    
    if (resetHeaderBgColorBtn) {
        resetHeaderBgColorBtn.addEventListener('click', () => {
            document.getElementById('headerBgColor').value = '#f8f9fa';
            applyFormattingToPreview();
        });
    }
    
    // Alignment button click handlers
    document.querySelectorAll('.alignment-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const button = e.currentTarget;
            const container = button.closest('.alignment-buttons');
            const alignment = button.getAttribute('data-align');
            
            // Update active state
            setAlignmentButtonActive(container.id, alignment);
            
            // Apply formatting to preview
            applyFormattingToPreview();
        });
    });
    
    // Format change handlers for live preview
    const controls = ['textColor', 'cellBgColor', 'headerBgColor'];
    
    controls.forEach(controlId => {
        const control = document.getElementById(controlId);
        if (control) {
            control.addEventListener('change', applyFormattingToPreview);
        }
    });
}

// Initialize formatting toolbar
function initializeFormattingButtons() {
    // Store reference to last focused input
    let lastFocusedInput = null;
    let lastSelectionStart = 0;
    let lastSelectionEnd = 0;
    
    // Track focus and selection on table inputs
    document.addEventListener('focus', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('#editableTable')) {
            lastFocusedInput = e.target;
        }
    }, true);
    
    document.addEventListener('select', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('#editableTable')) {
            lastSelectionStart = e.target.selectionStart;
            lastSelectionEnd = e.target.selectionEnd;
        }
    }, true);
    
    // Also track on mouseup and keyup for better selection detection
    document.addEventListener('mouseup', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('#editableTable')) {
            lastSelectionStart = e.target.selectionStart;
            lastSelectionEnd = e.target.selectionEnd;
        }
    });
    
    document.addEventListener('keyup', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('#editableTable')) {
            lastSelectionStart = e.target.selectionStart;
            lastSelectionEnd = e.target.selectionEnd;
        }
    });
    
    // Handle format button clicks
    document.addEventListener('click', (e) => {
        if (e.target.classList.contains('format-btn')) {
            e.preventDefault();
            const format = e.target.dataset.format;
            if (lastFocusedInput) {
                applyFormatting(format, lastFocusedInput, lastSelectionStart, lastSelectionEnd);
            }
        }
    });
}

// Apply formatting to selected text in input
function applyFormatting(format, input, start, end) {
    if (!input) return;
    
    // Focus back on the input
    input.focus();
    
    const text = input.value;
    
    if (start !== end) {
        // Text is selected
        const selectedText = text.substring(start, end);
        const beforeText = text.substring(0, start);
        const afterText = text.substring(end);
        
        let formattedText;
        if (format === 'subscript') {
            formattedText = `<sub>${selectedText}</sub>`;
        } else if (format === 'superscript') {
            formattedText = `<sup>${selectedText}</sup>`;
        }
        
        input.value = beforeText + formattedText + afterText;
        
        // Set cursor position after the inserted text
        const newPosition = start + formattedText.length;
        input.setSelectionRange(newPosition, newPosition);
        
        // Update visual preview
        updateCellPreview(input);
    }
}

// Update cell preview to show formatted text
function updateCellPreview(input) {
    // Create a preview div if it doesn't exist
    let preview = input.parentElement.querySelector('.format-preview');
    if (!preview) {
        preview = document.createElement('div');
        preview.className = 'format-preview';
        input.parentElement.appendChild(preview);
    }
    
    // Convert the input value to HTML with formatting
    const formattedHTML = input.value
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/&lt;sub&gt;(.*?)&lt;\/sub&gt;/g, '<sub>$1</sub>')
        .replace(/&lt;sup&gt;(.*?)&lt;\/sup&gt;/g, '<sup>$1</sup>');
    
    preview.innerHTML = formattedHTML;
    
    // Show/hide based on content
    if (formattedHTML.includes('<sub>') || formattedHTML.includes('<sup>')) {
        preview.style.display = 'block';
        input.style.display = 'none';
    } else {
        preview.style.display = 'none';
        input.style.display = 'block';
    }
}

// Add click handler to switch back to input when preview is clicked
document.addEventListener('click', (e) => {
    if (e.target.classList.contains('format-preview')) {
        const input = e.target.parentElement.querySelector('input');
        if (input) {
            e.target.style.display = 'none';
            input.style.display = 'block';
            input.focus();
        }
    }
});

function generateTablePreview() {
    const rows = parseInt(document.getElementById('tableRows').value);
    const cols = parseInt(document.getElementById('tableCols').value);
    const includeHeaders = document.getElementById('includeHeaders').checked;
    
    if (rows < 1 || cols < 1 || rows > 20 || cols > 10) {
        showMessage('Veuillez entrer des valeurs valides (1-20 lignes, 1-10 colonnes)', 'error');
        return;
    }
    
    const previewDiv = document.getElementById('tablePreview');
    const tableContainer = document.getElementById('editableTable');
    
    // Create editable table
    let tableHTML = '<table class="editable-table">';
    
    // Add header row if requested
    if (includeHeaders) {
        tableHTML += '<thead><tr>';
        for (let c = 0; c < cols; c++) {
            tableHTML += `<th><input type="text" placeholder="En-tête ${c + 1}" data-row="-1" data-col="${c}"></th>`;
        }
        tableHTML += '</tr></thead>';
    }
    
    // Add data rows
    tableHTML += '<tbody>';
    const startRow = includeHeaders ? 0 : 0;
    for (let r = startRow; r < rows; r++) {
        tableHTML += '<tr>';
        for (let c = 0; c < cols; c++) {
            tableHTML += `<td><input type="text" placeholder="Cellule ${r + 1},${c + 1}" data-row="${r}" data-col="${c}"></td>`;
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    
    tableContainer.innerHTML = tableHTML;
    previewDiv.style.display = 'block';
    
    // Add blur event listeners to update preview on all inputs
    const inputs = tableContainer.querySelectorAll('input');
    inputs.forEach(input => {
        input.addEventListener('blur', () => updateCellPreview(input));
        input.addEventListener('input', () => updateCellPreview(input));
    });
    
    // Apply formatting to the newly generated table
    applyFormattingToPreview();
}

// Apply selected formatting to table preview
function applyFormattingToPreview() {
    const table = document.querySelector('#editableTable .editable-table');
    if (!table) return;
    
    // Get all style values
    const textColor = document.getElementById('textColor').value;
    const textAlignment = getAlignmentValue('textAlignment');
    const verticalAlignment = getAlignmentValue('verticalAlignment');
    const cellBgColor = document.getElementById('cellBgColor').value;
    const headerBgColor = document.getElementById('headerBgColor').value;
    
    // Convert vertical alignment to CSS value
    const cssVerticalAlign = verticalAlignment === 'top' ? 'top' :
                           verticalAlignment === 'bottom' ? 'bottom' : 'middle';
    
    // Apply styles to header cells
    const headerCells = table.querySelectorAll('th');
    headerCells.forEach(cell => {
        cell.style.color = textColor;
        cell.style.textAlign = textAlignment;
        cell.style.verticalAlign = cssVerticalAlign;
        cell.style.backgroundColor = headerBgColor;
        
        // Also apply to inputs and previews inside cells
        const input = cell.querySelector('input');
        const preview = cell.querySelector('.format-preview');
        if (input) {
            input.style.color = textColor;
            input.style.textAlign = textAlignment;
            input.style.backgroundColor = 'transparent';
        }
        if (preview) {
            preview.style.color = textColor;
            preview.style.textAlign = textAlignment;
        }
    });
    
    // Apply styles to data cells
    const dataCells = table.querySelectorAll('td');
    dataCells.forEach(cell => {
        cell.style.color = textColor;
        cell.style.textAlign = textAlignment;
        cell.style.verticalAlign = cssVerticalAlign;
        cell.style.backgroundColor = cellBgColor;
        
        // Also apply to inputs and previews inside cells
        const input = cell.querySelector('input');
        const preview = cell.querySelector('.format-preview');
        if (input) {
            input.style.color = textColor;
            input.style.textAlign = textAlignment;
            input.style.backgroundColor = 'transparent';
        }
        if (preview) {
            preview.style.color = textColor;
            preview.style.textAlign = textAlignment;
        }
    });
}

// Parse text with subscript/superscript tags and insert formatted text into Word
async function insertFormattedText(cell, text) {
    // Handle empty text
    if (!text || text.trim() === '') {
        cell.body.insertText('', "Start");
        return;
    }
    
    // Regular expression to find <sub> and <sup> tags
    const regex = /<(sub|sup)>(.*?)<\/\1>/g;
    let lastIndex = 0;
    let match;
    let hasFormatting = false;
    
    while ((match = regex.exec(text)) !== null) {
        hasFormatting = true;
        
        // Insert text before the tag
        if (match.index > lastIndex) {
            const plainText = text.substring(lastIndex, match.index);
            cell.body.insertText(plainText, "End");
        }
        
        // Insert formatted text
        const tagType = match[1];
        const tagContent = match[2];
        const formattedText = cell.body.insertText(tagContent, "End");
        
        if (tagType === 'sub') {
            formattedText.font.subscript = true;
        } else if (tagType === 'sup') {
            formattedText.font.superscript = true;
        }
        
        lastIndex = regex.lastIndex;
    }
    
    // Insert remaining text after the last tag, or all text if no formatting
    if (lastIndex < text.length || !hasFormatting) {
        const remainingText = hasFormatting ? text.substring(lastIndex) : text;
        cell.body.insertText(remainingText, "End");
    }
}

async function insertTableToWord() {
    try {
        const rows = parseInt(document.getElementById('tableRows').value);
        const cols = parseInt(document.getElementById('tableCols').value);
        const includeHeaders = document.getElementById('includeHeaders').checked;
        
        // Validate input
        if (isNaN(rows) || isNaN(cols) || rows < 1 || cols < 1) {
            showMessage('Valeurs de tableau invalides', 'error');
            return;
        }
        
        // Collect table data from inputs
        const inputs = document.querySelectorAll('#editableTable input');
        const dataMap = new Map();
        inputs.forEach(input => {
            const row = parseInt(input.dataset.row);
            const col = parseInt(input.dataset.col);
            const key = `${row}-${col}`;
            // Always get value from input, even if preview is showing
            dataMap.set(key, input.value || '');
        });
        
        await Word.run(async (context) => {
            // Get the document body instead of selection
            const body = context.document.body;
            
            // Insert a paragraph first to ensure we have a valid location
            const paragraph = body.insertParagraph("", "End");
            
            // Insert table at the paragraph location
            const actualRows = includeHeaders ? rows + 1 : rows;
            const table = paragraph.insertTable(actualRows, cols, "After");
            
            // Sync to ensure table is created
            await context.sync();
            
            // Get selected styles
            const textColor = document.getElementById('textColor').value;
            const textAlignment = getAlignmentValue('textAlignment');
            const verticalAlignment = getAlignmentValue('verticalAlignment');
            const cellBgColor = document.getElementById('cellBgColor').value;
            const headerBgColor = document.getElementById('headerBgColor').value;
            
            // Now fill the table cells one by one
            for (let r = 0; r < actualRows; r++) {
                for (let c = 0; c < cols; c++) {
                    try {
                        const cell = table.getCell(r, c);
                        let cellValue = '';
                        
                        if (r === 0 && includeHeaders) {
                            // Header row
                            cellValue = dataMap.get(`-1-${c}`) || `En-tête ${c + 1}`;
                        } else {
                            // Data rows
                            const dataRow = includeHeaders ? r - 1 : r;
                            cellValue = dataMap.get(`${dataRow}-${c}`) || '';
                        }
                        
                        // Insert formatted text first
                        await insertFormattedText(cell, cellValue);
                        
                        // Apply cell styling after text insertion
                        cell.body.font.color = textColor;
                        
                        // Set vertical alignment
                        const wordVerticalAlign = verticalAlignment === 'top' ? 'Top' :
                                                verticalAlignment === 'bottom' ? 'Bottom' : 'Center';
                        cell.verticalAlignment = wordVerticalAlign;
                        
                        // Set horizontal alignment
                        const wordHorizontalAlign = textAlignment === 'left' ? 'Left' : 
                                                  textAlignment === 'center' ? 'Centered' : 'Right';
                        cell.horizontalAlignment = wordHorizontalAlign;
                        
                        // Set cell background color using shadingColor
                        if (r === 0 && includeHeaders) {
                            cell.shadingColor = headerBgColor;
                            cell.body.font.bold = true;
                        } else {
                            cell.shadingColor = cellBgColor;
                        }
                        
                        // Set cell padding for better spacing
                        cell.setCellPadding("Top", 6);    // 6 points top padding
                        cell.setCellPadding("Bottom", 6); // 6 points bottom padding
                        cell.setCellPadding("Left", 8);   // 8 points left padding
                        cell.setCellPadding("Right", 8);  // 8 points right padding
                        
                        // Sync after each cell to avoid issues
                        await context.sync();
                    } catch (cellError) {
                        console.error(`Error in cell ${r},${c}:`, cellError);
                    }
                }
            }
            
            // Remove the empty paragraph we created
            paragraph.delete();
            
            // Final sync
            await context.sync();
        });
        
        showMessage('Tableau inséré avec succès!', 'success');
        resetTableGenerator();
        
    } catch (error) {
        console.error('Error inserting table:', error);
        console.error('Error details:', error.debugInfo);
        showMessage('Erreur lors de l\'insertion du tableau. Veuillez réessayer.', 'error');
    }
}


function resetTableGenerator() {
    document.getElementById('tableRows').value = '3';
    document.getElementById('tableCols').value = '3';
    document.getElementById('includeHeaders').checked = true;
    document.getElementById('tablePreview').style.display = 'none';
    document.getElementById('editableTable').innerHTML = '';
    
    // Reset colors to saved preferences
    applyPreferencesToUI();
}

async function insertTitlePage(event) {
    event.preventDefault();
    
    const codeDuCours = document.getElementById("title").value;
    const section = document.getElementById("labName").value;
    const numeroEtNomLab = document.getElementById("professorName").value;
    const nomAssistant = document.getElementById("date").value;
    const actualDate = document.getElementById("actualDate").value;
    const workSpace = document.getElementById("workSpace").value;
    const studentInfo = document.getElementById("studentInfo").value;
    
    // Format the date to be more readable
    const formattedDate = new Date(actualDate).toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    
    try {
        await Word.run(async (context) => {
            // Get the document body
            const body = context.document.body;
            
            // Clear the document and start fresh
            body.clear();
            
            // Insert Code du cours as the first paragraph
            let currentParagraph = body.insertParagraph(codeDuCours, "Start");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after Code du cours (3 empty lines to fit 7 fields on one page)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert Section
            currentParagraph = currentParagraph.insertParagraph(section, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after Section (4 empty lines)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert # et nom du laboratoire
            currentParagraph = currentParagraph.insertParagraph(numeroEtNomLab, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after # et nom du laboratoire (4 empty lines)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert Nom de l'assistant
            currentParagraph = currentParagraph.insertParagraph(nomAssistant, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after Nom de l'assistant (4 empty lines)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert Date
            currentParagraph = currentParagraph.insertParagraph(formattedDate, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after Date (4 empty lines)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert Espace de travail
            currentParagraph = currentParagraph.insertParagraph(workSpace, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Add spacing after Espace de travail (4 empty lines)
            for (let i = 0; i < 3; i++) {
                currentParagraph = currentParagraph.insertParagraph("", "After");
            }
            
            // Insert Nom et no. etudiant
            currentParagraph = currentParagraph.insertParagraph(studentInfo, "After");
            currentParagraph.set({
                alignment: "Centered",
                style: "Normal"
            });
            currentParagraph.font.set({
                name: "Times New Roman",
                size: 12
            });
            
            // Sync after all formatting is applied
            await context.sync();
            
            // Load all paragraphs to ensure formatting is applied
            body.load('paragraphs');
            await context.sync();
            
            // Re-apply formatting to all non-empty paragraphs to ensure it sticks
            for (let i = 0; i < body.paragraphs.items.length; i++) {
                const para = body.paragraphs.items[i];
                para.load('text');
                await context.sync();
                
                if (para.text && para.text.trim() !== '') {
                    para.set({
                        alignment: "Centered",
                        style: "Normal"
                    });
                    para.font.set({
                        name: "Times New Roman",
                        size: 12
                    });
                }
            }
            
            await context.sync();
        });
        
        // Clear the form after successful insertion
        document.getElementById("titlePageForm").reset();
        
        // Show success message
        showMessage("Title page inserted successfully!", "success");
        
    } catch (error) {
        console.error('Error inserting title page:', error);
        showMessage("Error inserting title page. Please try again.", "error");
    }
}

function showMessage(text, type) {
    // Remove any existing message
    const existingMessage = document.querySelector('.message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // Create and show new message
    const message = document.createElement('div');
    message.className = `message ${type}`;
    message.textContent = text;
    
    const container = document.querySelector('.container');
    container.insertBefore(message, container.firstChild);
    
    // Remove message after 3 seconds
    setTimeout(() => {
        if (message.parentNode) {
            message.remove();
        }
    }, 3000);
}