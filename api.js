class GeminiAPI {
    constructor() {
        this.apiKey = localStorage.getItem('GEMINI_API_KEY') || '';
    }

    setApiKey(key) {
        this.apiKey = key.trim();
        localStorage.setItem('GEMINI_API_KEY', this.apiKey);
    }

    getApiKey() {
        return this.apiKey;
    }

    async generateContent(prompt, systemInstruction = "") {
        if (!this.apiKey) {
            throw new Error("API Key is missing. Click the ⚙️ icon to set your Google Gemini API Key.");
        }

        // Direct call to Gemini Native REST Endpoint
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${this.apiKey}`;

        const payload = {
            contents: [
                {
                    parts: [
                        { text: systemInstruction ? `[Instruction: ${systemInstruction}]\n\n${prompt}` : prompt }
                    ]
                }
            ]
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error?.message || `API Request Failed with status ${response.status}`);
        }

        if (data.candidates && data.candidates[0]?.content?.parts[0]?.text) {
            return data.candidates[0].content.parts[0].text;
        } else {
            throw new Error("Received an unexpected response structure from Gemini API.");
        }
    }
}

const geminiService = new GeminiAPI();