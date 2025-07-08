// Backup of original insertTableToWord function
async function insertTableToWord() {
    try {
        const rows = parseInt(document.getElementById('tableRows').value);
        const cols = parseInt(document.getElementById('tableCols').value);
        const includeHeaders = document.getElementById('includeHeaders').checked;
        
        // Collect table data from inputs
        const inputs = document.querySelectorAll('#editableTable input');
        
        // Organize data by row and column
        const dataMap = new Map();
        inputs.forEach(input => {
            const row = parseInt(input.dataset.row);
            const col = parseInt(input.dataset.col);
            const key = `${row}-${col}`;
            dataMap.set(key, input.value || '');
        });
        
        await Word.run(async (context) => {
            // Get the current selection (cursor position)
            const selection = context.document.getSelection();
            
            // Insert table at cursor position
            const actualRows = includeHeaders ? rows + 1 : rows;
            const table = selection.insertTable(actualRows, cols, "Replace", []);
            
            // Load the table before setting properties
            context.load(table);
            await context.sync();
            
            // Set table properties after loading
            table.horizontalAlignment = "Centered";
            
            // Sync after setting properties
            await context.sync();
            
            // Fill header row if included
            if (includeHeaders) {
                for (let c = 0; c < cols; c++) {
                    const cellValue = dataMap.get(`-1-${c}`) || `En-tête ${c + 1}`;
                    const cell = table.getCell(0, c);
                    cell.body.insertText(cellValue, "Replace");
                    // Set header formatting
                    const paragraph = cell.body.paragraphs.getFirst();
                    paragraph.font.bold = true;
                }
            }
            
            // Fill data rows
            const startRow = includeHeaders ? 1 : 0;
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cellValue = dataMap.get(`${r}-${c}`) || '';
                    const cell = table.getCell(startRow + r, c);
                    cell.body.insertText(cellValue, "Replace");
                }
            }
            
            // Final sync
            await context.sync();
        });
        
        showMessage('Tableau inséré avec succès!', 'success');
        resetTableGenerator();
        
    } catch (error) {
        console.error('Error inserting table:', error);
        showMessage('Erreur lors de l\'insertion du tableau. Veuillez réessayer.', 'error');
    }
}