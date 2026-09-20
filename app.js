Office.onReady((info) => {
    const hostBadge = document.getElementById('host-badge');
    const quickActionsContainer = document.getElementById('quick-actions');
    const apiKeyInput = document.getElementById('api-key-input');

    if (hostBadge) hostBadge.innerText = info.host || "Web/Standalone";
    if (apiKeyInput) apiKeyInput.value = geminiService.getApiKey();

    // Load Host-Specific Quick Action Buttons
    loadQuickActions(info.host, quickActionsContainer);

    // Event Listeners
    document.getElementById('toggle-settings')?.addEventListener('click', () => {
        document.getElementById('settings-panel').classList.toggle('hidden');
    });

    document.getElementById('save-key-btn')?.addEventListener('click', () => {
        const key = apiKeyInput.value;
        geminiService.setApiKey(key);
        const status = document.getElementById('key-status');
        if (status) {
            status.innerText = "Saved!";
            status.style.color = "green";
            setTimeout(() => { status.innerText = ""; }, 2000);
        }
    });

    document.getElementById('send-btn')?.addEventListener('click', handleUserSend);
    document.getElementById('user-input')?.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && !event.shiftKey) {
            event.preventDefault();
            handleUserSend();
        }
    });
});

async function handleUserSend() {
    const inputEl = document.getElementById('user-input');
    const text = inputEl.value.trim();
    if (!text) return;

    appendMessage(text, 'user-msg');
    inputEl.value = '';

    const loadingEl = appendMessage("Thinking...", 'ai-msg');

    try {
        let response = "";
        
        // Regex pattern to check if the prompt is asking for image creation
        const isImageRequest = /\b(image|picture|draw|photo|illustration|generate an image|create an image|make an image)\b/i.test(text);

        if (isImageRequest) {
            // Call Imagen 3 API via GeminiService
            response = await geminiService.generateImage(text);
            loadingEl.innerHTML = response; // Render base64 <img> tag directly
        } else {
            // System prompt engineered to force raw content generation instead of step-by-step instructions
            const systemPrompt = 
                "You are a direct content generator inside Microsoft Office. " +
                "CRITICAL INSTRUCTION: Do NOT provide meta-instructions, step-by-step guides, or advice on how to use MS Word/Excel features. " +
                "Directly output the final document content, text, tables, worksheets, or code as requested so it can be inserted into the document.";

            response = await geminiService.generateContent(text, systemPrompt);
            loadingEl.innerText = response;
        }

        // Add "Insert into Document" action button under AI responses
        const insertBtn = document.createElement('button');
        insertBtn.className = 'btn-chip';
        insertBtn.style.marginTop = '8px';
        insertBtn.innerText = '📥 Insert into Document';
        insertBtn.onclick = async () => {
            try {
                await OfficeIntegration.insertText(response);
            } catch (err) {
                alert(`Insertion failed: ${err.message}`);
            }
        };

        loadingEl.appendChild(document.createElement('br'));
        loadingEl.appendChild(insertBtn);

    } catch (err) {
        loadingEl.innerText = `Error: ${err.message}`;
        loadingEl.className = 'message system-msg';
    }
}

function appendMessage(text, className) {
    const container = document.getElementById('chat-container');
    const msg = document.createElement('div');
    msg.className = `message ${className}`;
    msg.innerText = text;
    container.appendChild(msg);
    container.scrollTop = container.scrollHeight;
    return msg;
}

function loadQuickActions(host, container) {
    if (!container) return;
    container.innerHTML = '';
    let actions = [];

    if (host === Office.HostType.Word) {
        actions = [
            { label: '✨ Reformat Selection', prompt: 'Reformat the selected text to be professional, clear, and well-structured. Output only the reformatted text.' },
            { label: '📝 Summarize Doc', prompt: 'Provide a concise bullet-point summary of the following text.' },
            { label: '🎨 Create Image', prompt: 'Create an image illustration of ' },
            { label: '📐 Create Worksheet', prompt: 'Generate a printable worksheet with questions and spacing.' }
        ];
    } else if (host === Office.HostType.Excel) {
        actions = [
            { label: '📊 Create Formula', prompt: 'Write an Excel formula for the following requirement. Return only the formula.' },
            { label: '🧹 Clean Cell Data', prompt: 'Clean and standardize the selected cell data values.' }
        ];
    } else {
        actions = [
            { label: '💡 Summarize Context', prompt: 'Summarize the selected content.' },
            { label: '✉️ Draft Response', prompt: 'Draft a polite and professional response based on this text.' }
        ];
    }

    actions.forEach(act => {
        const btn = document.createElement('button');
        btn.className = 'btn-chip';
        btn.innerText = act.label;
        btn.onclick = async () => {
            const contextData = await OfficeIntegration.readSelection();
            const fullPrompt = contextData 
                ? `${act.prompt}\n\nSelected Data/Text:\n${contextData}`
                : act.prompt;
            document.getElementById('user-input').value = fullPrompt;
            handleUserSend();
        };
        container.appendChild(btn);
    });
}