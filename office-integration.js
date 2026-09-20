const OfficeIntegration = {
    getHost() {
        if (!window.Office || !Office.context) return "Unknown";
        return Office.context.host; // Returns 'Word', 'Excel', 'PowerPoint', 'Outlook', or 'OneNote'
    },

    async readSelection() {
        const host = this.getHost();

        if (host === Office.HostType.Word) {
            return await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.load("text");
                await context.sync();
                return range.text;
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
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                    resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
                });
            });
        }
    },

    async insertText(text) {
        const host = this.getHost();

        if (host === Office.HostType.Word) {
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.insertText(text, Word.InsertLocation.replace);
                await context.sync();
            });
        } 
        else if (host === Office.HostType.Excel) {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                // Clean potential backticks if AI outputs code blocks
                const cleanText = text.replace(/```[a-z]*\n?/gi, '').replace(/```/g, '').trim();
                range.values = [[cleanText]];
                await context.sync();
            });
        } 
        else {
            Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
        }
    }
};