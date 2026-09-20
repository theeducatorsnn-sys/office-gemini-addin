const OfficeIntegration = {
    getHost() {
        if (!window.Office || !Office.context) return "Unknown";
        return Office.context.host; // Returns 'Word', 'Excel', 'PowerPoint', 'Outlook', or 'OneNote'
    },

    async readSelection() {
        const host = this.getHost();

        try {
            if (host === Office.HostType.Word) {
                return await Word.run(async (context) => {
                    const range = context.document.getSelection();
                    range.load("text");
                    await context.sync();
                    return range.text || "";
                });
            } 
            else if (host === Office.HostType.Excel) {
                return await Excel.run(async (context) => {
                    const range = context.workbook.getSelectedRange();
                    range.load("values");
                    await context.sync();
                    return JSON.stringify(range.values);
                });
            } 
            else {
                return new Promise((resolve) => {
                    if (!Office.context.document || !Office.context.document.getSelectedDataAsync) {
                        resolve("");
                        return;
                    }
                    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
                    });
                });
            }
        } catch (err) {
            console.warn("Could not read selection:", err);
            return "";
        }
    },

    async insertText(text) {
        if (!text) return;
        const host = this.getHost();

        // 1. Word Host Handling
        if (host === Office.HostType.Word) {
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                
                // Convert Markdown to HTML for styled insertion into Word
                const formattedHtml = this._markdownToHtml(text);
                
                try {
                    // Try inserting formatted HTML so headings, bold, and lists render properly
                    range.insertHtml(formattedHtml, Word.InsertLocation.replace);
                } catch (e) {
                    // Fallback to standard text insertion if HTML parsing fails
                    range.insertText(text, Word.InsertLocation.replace);
                }
                
                await context.sync();
            });
        } 
        // 2. Excel Host Handling
        else if (host === Office.HostType.Excel) {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                
                // Clean potential Markdown code block ticks
                const cleanText = text.replace(/```[a-z]*\n?/gi, '').replace(/```/g, '').trim();

                // Attempt to parse text as 2D Array or CSV for multi-cell insertion
                const parsedGrid = this._parseTextToGrid(cleanText);

                if (parsedGrid && parsedGrid.length > 0) {
                    // Resize selected range to match grid dimensions
                    const targetRange = range.getResizedRange(
                        parsedGrid.length - 1, 
                        parsedGrid[0].length - 1
                    );
                    targetRange.values = parsedGrid;
                } else {
                    range.values = [[cleanText]];
                }

                await context.sync();
            });
        } 
        // 3. Fallback for PowerPoint, Outlook, OneNote
        else {
            return new Promise((resolve, reject) => {
                if (!Office.context.document || !Office.context.document.setSelectedDataAsync) {
                    reject(new Error("Document context unavailable."));
                    return;
                }
                Office.context.document.setSelectedDataAsync(
                    text, 
                    { coercionType: Office.CoercionType.Text }, 
                    (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve();
                        } else {
                            reject(new Error(result.error ? result.error.message : "Failed to insert text."));
                        }
                    }
                );
            });
        }
    },

    // Internal Helper: Convert Markdown strings into HTML for MS Word
    _markdownToHtml(markdown) {
        let html = markdown
            // Replace Headers
            .replace(/^### (.*$)/gim, '<h3>$1</h3>')
            .replace(/^## (.*$)/gim, '<h2>$1</h2>')
            .replace(/^# (.*$)/gim, '<h1>$1</h1>')
            // Replace Bold and Italics
            .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
            .replace(/\*(.*?)\*/g, '<i>$1</i>')
            // Replace Unordered Lists
            .replace(/^\* (.*$)/gim, '<ul><li>$1</li></ul>')
            .replace(/^- (.*$)/gim, '<ul><li>$1</li></ul>')
            // Clean duplicate list tags
            .replace(/<\/ul>\s*<ul>/g, '')
            // Line Breaks
            .replace(/\n/g, '<br/>');

        return `<div>${html}</div>`;
    },

    // Internal Helper: Parse structured strings/JSON into Excel 2D cell grids
    _parseTextToGrid(text) {
        try {
            // Check if AI output is a JSON 2D Array (e.g. [["Name", "Age"], ["Ali", 25]])
            const jsonData = JSON.parse(text);
            if (Array.isArray(jsonData) && Array.isArray(jsonData[0])) {
                return jsonData;
            }
        } catch (e) {
            // Not JSON, continue to CSV/TSV parsing
        }

        // CSV/Tab-separated multi-line text fallback
        if (text.includes('\t') || text.includes(',')) {
            const lines = text.split('\n').filter(line => line.trim().length > 0);
            if (lines.length > 1) {
                const delimiter = text.includes('\t') ? '\t' : ',';
                return lines.map(line => line.split(delimiter).map(cell => cell.trim()));
            }
        }

        return null;
    }
};