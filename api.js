class GeminiAPI {
    constructor() {
        this.apiKey = localStorage.getItem('GEMINI_API_KEY') || 'AQ.Ab8RN6Jwr83w-Wlf4Ikg3euUf4YAVJsQR6u1ePIKH9YmeNTVBw';
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

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(this.apiKey)}`;

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

    async generateImage(prompt) {
        if (!this.apiKey) {
            throw new Error("API Key is missing. Click the ⚙️ icon to set your Google Gemini API Key.");
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key=${encodeURIComponent(this.apiKey)}`;

        const payload = {
            instances: [
                { prompt: prompt }
            ],
            parameters: {
                sampleCount: 1,
                aspectRatio: "1:1"
            }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error?.message || `Image generation failed with status ${response.status}`);
        }

        if (data.predictions && data.predictions[0]?.bytesBase64Encoded) {
            const base64Image = data.predictions[0].bytesBase64Encoded;
            return `<img src="data:image/png;base64,${base64Image}" style="max-width:350px; height:auto; display:block; margin:10px 0;" />`;
        }

        throw new Error("No image data returned from Imagen API.");
    }
}

const geminiService = new GeminiAPI();