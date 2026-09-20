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

        // Direct call to Gemini REST Endpoint
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${this.apiKey}`;

        // Construct standard payload with system_instruction support
        const payload = {
            contents: [
                {
                    role: "user",
                    parts: [{ text: prompt }]
                }
            ]
        };

        if (systemInstruction) {
            payload.system_instruction = {
                parts: [{ text: systemInstruction }]
            };
        }

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error?.message || `API Request Failed with status ${response.status}`);
        }

        // Check if response was blocked or incomplete
        const candidate = data.candidates && data.candidates[0];
        if (candidate) {
            if (candidate.finishReason === "SAFETY") {
                throw new Error("Response was flagged and blocked due to safety settings.");
            }
            
            if (candidate.content && candidate.content.parts && candidate.content.parts[0]?.text) {
                return candidate.content.parts[0].text;
            }
        }

        throw new Error("Received an unexpected or empty response structure from Gemini API.");
    }
}

const geminiService = new GeminiAPI();