Office.onReady((info) => {
    const hostBadge = document.getElementById('host-badge');
    const quickActionsContainer = document.getElementById('quick-actions');
    const apiKeyInput = document.getElementById('api-key-input');

    hostBadge.innerText = info.host || "Web/Standalone";
    apiKeyInput.value = geminiService.getApiKey();

    // Load Host-Specific Quick Action Buttons
    loadQuickActions(info.host, quickActionsContainer);

    // Event Listeners
    document.getElementById('toggle-settings').addEventListener('click', () => {
        document.getElementById('settings-panel').classList.toggle('hidden');
    });

    document.getElementById('save-key-btn').addEventListener('click', () => {
        const key = apiKeyInput.value;
        geminiService.setApiKey(key);
        const status = document.getElementById('key-status');
        status.innerText = "Saved!";
        status.style.color = "green";
        setTimeout(() => { status.innerText = ""; }, 2000);
    });

    document.getElementById('send-btn').addEventListener('click', handleUserSend);
    document.getElementById('user-input').addEventListener('keydown', (event) => {
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

    try {
        const loadingEl = appendMessage("Thinking...", 'ai-msg');
        const response = await geminiService.generateContent(text, "You are a helpful MS Office assistant.");
        loadingEl.innerText = response;

        // Add an "Insert into Document" action button under AI responses
        const insertBtn = document.createElement('button');
        insertBtn.className = 'btn-chip';
        insertBtn.style.marginTop = '6px';
        insertBtn.innerText = '📥 Insert into Document';
        insertBtn.onclick = () => OfficeIntegration.insertText(response);
        loadingEl.appendChild(document.createElement('br'));
        loadingEl.appendChild(insertBtn);

    } catch (err) {
        appendMessage(`Error: ${err.message}`, 'system-msg');
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
    container.innerHTML = '';
    let actions = [];

    if (host === Office.HostType.Word) {
        actions = [
            { label: '✨ Reformat Selection', prompt: 'Reformat the selected text to be professional, clear, and well-structured.' },
            { label: '📝 Summarize Doc', prompt: 'Provide a concise bullet-point summary of the following text.' }
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
            const fullPrompt = `${act.prompt}\n\nSelected Data/Text:\n${contextData}`;
            document.getElementById('user-input').value = fullPrompt;
            handleUserSend();
        };
        container.appendChild(btn);
    });
}