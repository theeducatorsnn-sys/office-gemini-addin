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

    /**
     * Generates text content using the gemini-2.5-flash model.
     */
    async generateContent(prompt, systemInstruction = "") {
        if (!this.apiKey) {
            throw new Error("API Key is missing. Click the ⚙️ icon in the add-in sidebar to paste your Google Gemini API Key.");
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(this.apiKey)}`;

        const payload = {
            contents: [{ role: "user", parts: [{ text: prompt }] }]
        };

        if (systemInstruction) {
            payload.system_instruction = { parts: [{ text: systemInstruction }] };
        }

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.error?.message || `API Request Failed: ${response.status}`);
        }

        const candidate = data.candidates && data.candidates[0];
        
        if (candidate) {
            if (candidate.finishReason === "SAFETY") {
                throw new Error("Response was flagged and blocked due to Google safety settings.");
            }

            if (candidate.content && candidate.content.parts && candidate.content.parts[0]?.text) {
                return candidate.content.parts[0].text;
            }
        }

        throw new Error("Invalid or empty response structure received from Gemini API.");
    }

    /**
     * Generates an AI image using Google's Imagen 3 model and returns an HTML <img> tag with Base64 data.
     */
    async generateImage(prompt) {
        if (!this.apiKey) {
            throw new Error("API Key is missing. Click the ⚙️ icon in the add-in sidebar to paste your Google Gemini API Key.");
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key=${encodeURIComponent(this.apiKey)}`;

        const payload = {
            instances: [{ prompt: prompt }],
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
            throw new Error(data.error?.message || `Image generation failed with status code ${response.status}`);
        }

        if (data.predictions && data.predictions[0]?.bytesBase64Encoded) {
            const base64Image = data.predictions[0].bytesBase64Encoded;
            return `<img src="data:image/png;base64,${base64Image}" style="max-width:300px; height:auto; display:block; margin:10px 0; border-radius: 6px;" />`;
        }

        throw new Error("No image data was returned. The prompt may have triggered content safety filters.");
    }
}

const geminiService = new GeminiAPI();