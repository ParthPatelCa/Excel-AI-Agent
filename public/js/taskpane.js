Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('ExcelAI Assistant loaded successfully');
        initializeTaskPane();
    }
});

let selectedRange = null;
let chatHistory = [];

function initializeTaskPane() {
    const sendBtn = document.getElementById('sendBtn');
    const chatInput = document.getElementById('chatInput');
    const refreshDataBtn = document.getElementById('refreshDataBtn');
    const clearChatBtn = document.getElementById('clearChatBtn');
    const charCounter = document.getElementById('charCounter');

    // Event listeners
    sendBtn.addEventListener('click', handleSendMessage);
    chatInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSendMessage();
        }
    });
    chatInput.addEventListener('input', updateInputState);
    refreshDataBtn.addEventListener('click', refreshSelectedData);
    clearChatBtn.addEventListener('click', clearChat);

    // Quick action buttons
    document.querySelectorAll('.action-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const action = e.currentTarget.dataset.action;
            handleQuickAction(action);
        });
    });

    // Initialize
    refreshSelectedData();
    updateInputState();
}

async function refreshSelectedData() {
    updateStatus('Refreshing data...', 'loading');
    
    try {
        await Excel.run(async (context) => {
            const selection = context.workbook.getSelectedRange();
            selection.load(['address', 'values', 'rowCount', 'columnCount']);
            
            await context.sync();
            
            selectedRange = {
                address: selection.address,
                values: selection.values,
                rowCount: selection.rowCount,
                columnCount: selection.columnCount
            };
            
            updateDataContext();
            updateInputState();
            updateStatus('Ready', 'ready');
        });
    } catch (error) {
        console.error('Error refreshing data:', error);
        showError('Failed to refresh data. Please select a range in Excel and try again.');
        updateStatus('Error', 'error');
    }
}

function updateDataContext() {
    const contextContent = document.getElementById('contextContent');
    
    if (selectedRange && selectedRange.values && selectedRange.values.length > 0) {
        const sampleData = selectedRange.values.slice(0, 3);
        const hasNumbers = selectedRange.values.some(row => 
            row.some(cell => typeof cell === 'number')
        );
        const hasText = selectedRange.values.some(row => 
            row.some(cell => typeof cell === 'string')
        );
        
        contextContent.innerHTML = `
            <div class="data-info">
                <p><strong>Range:</strong> ${selectedRange.address}</p>
                <p><strong>Size:</strong> ${selectedRange.rowCount} rows × ${selectedRange.columnCount} columns</p>
                <p><strong>Data Types:</strong> ${hasNumbers ? 'Numbers' : ''} ${hasNumbers && hasText ? '& ' : ''}${hasText ? 'Text' : ''}</p>
                <details style="margin-top: 0.5rem;">
                    <summary style="cursor: pointer; color: #0078d4;">Sample Data (first 3 rows)</summary>
                    <pre style="background: #f8f9fa; padding: 0.5rem; border-radius: 4px; margin-top: 0.5rem; font-size: 0.8rem; overflow-x: auto;">${JSON.stringify(sampleData, null, 2)}</pre>
                </details>
            </div>
        `;
    } else {
        contextContent.innerHTML = '<p class="no-selection">No data selected. Please select a range in Excel.</p>';
    }
}

function updateInputState() {
    const sendBtn = document.getElementById('sendBtn');
    const chatInput = document.getElementById('chatInput');
    const charCounter = document.getElementById('charCounter');
    
    const messageLength = chatInput.value.length;
    const hasText = messageLength > 0;
    const hasRange = selectedRange !== null && selectedRange.values && selectedRange.values.length > 0;
    
    // Update character counter
    charCounter.textContent = `${messageLength}/2000`;
    if (messageLength > 1800) {
        charCounter.style.color = '#f44336';
    } else if (messageLength > 1500) {
        charCounter.style.color = '#ff9800';
    } else {
        charCounter.style.color = '#6c757d';
    }
    
    // Update send button state
    sendBtn.disabled = !hasText || !hasRange || messageLength > 2000;
}

function updateStatus(text, type) {
    const statusText = document.querySelector('.status-text');
    const statusDot = document.querySelector('.status-dot');
    
    statusText.textContent = text;
    
    // Remove all status classes
    statusDot.className = 'status-dot';
    
    // Add appropriate class
    switch (type) {
        case 'loading':
            statusDot.style.background = '#ff9800';
            break;
        case 'error':
            statusDot.style.background = '#f44336';
            break;
        case 'ready':
        default:
            statusDot.style.background = '#4caf50';
            break;
    }
}

async function handleQuickAction(action) {
    if (!selectedRange || !selectedRange.values || selectedRange.values.length === 0) {
        showError('Please select a data range first.');
        return;
    }

    const prompts = {
        analyze: 'Analyze this data and provide key insights, trends, and statistics.',
        summarize: 'Generate a summary of this data including count, averages, and key statistics.',
        chart: 'What type of chart would be best for visualizing this data? Explain why.',
        formulas: 'Suggest useful Excel formulas I could use with this data.'
    };

    const message = prompts[action];
    if (message) {
        document.getElementById('chatInput').value = message;
        updateInputState();
        await handleSendMessage();
    }
}

async function handleSendMessage() {
    const chatInput = document.getElementById('chatInput');
    const sendBtn = document.getElementById('sendBtn');
    const message = chatInput.value.trim();
    
    if (!message || !selectedRange) return;
    
    // Show loading state
    showLoadingState(true);
    addMessage(message, 'user');
    chatInput.value = '';
    updateInputState();
    
    try {
        const response = await fetch('/api/chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                message: message,
                selectedData: selectedRange,
                chatHistory: chatHistory.slice(-5) // Last 5 messages for context
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            addMessage(data.response, 'ai');
            
            // Handle any actions returned by the AI
            if (data.action) {
                addActionButton(data.action);
            }
        } else {
            addMessage(data.error || 'Sorry, something went wrong. Please try again.', 'error');
        }
        
    } catch (error) {
        console.error('Error sending message:', error);
        addMessage('Failed to connect to AI service. Please check your connection and try again.', 'error');
    } finally {
        showLoadingState(false);
    }
}

function showLoadingState(isLoading) {
    const sendBtn = document.getElementById('sendBtn');
    const btnText = sendBtn.querySelector('.btn-text');
    const btnLoading = sendBtn.querySelector('.btn-loading');
    
    if (isLoading) {
        btnText.style.display = 'none';
        btnLoading.style.display = 'flex';
        sendBtn.disabled = true;
        updateStatus('Processing...', 'loading');
    } else {
        btnText.style.display = 'block';
        btnLoading.style.display = 'none';
        updateInputState(); // This will properly set the disabled state
        updateStatus('Ready', 'ready');
    }
}

function addMessage(message, type) {
    const chatContainer = document.getElementById('chatContainer');
    
    // Remove welcome message if it exists and this is the first user message
    const welcomeMessage = chatContainer.querySelector('.welcome-message');
    if (welcomeMessage && type === 'user') {
        welcomeMessage.remove();
    }
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${type}-message`;
    
    // Handle multiline messages
    const formattedMessage = message.replace(/\n/g, '<br>');
    messageDiv.innerHTML = formattedMessage;
    
    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
    
    // Add to history
    chatHistory.push({ type, message, timestamp: Date.now() });
    
    // Keep only last 20 messages
    if (chatHistory.length > 20) {
        chatHistory = chatHistory.slice(-20);
    }
}

function addActionButton(action) {
    const chatContainer = document.getElementById('chatContainer');
    const actionDiv = document.createElement('div');
    actionDiv.className = 'message ai-message';
    
    const button = document.createElement('button');
    button.textContent = action.buttonText || 'Insert into Excel';
    button.className = 'btn-primary';
    button.style.marginTop = '0.5rem';
    
    button.addEventListener('click', () => executeAction(action));
    
    const description = document.createElement('p');
    description.textContent = action.description || 'Click to insert this into Excel';
    description.style.marginBottom = '0.5rem';
    
    actionDiv.appendChild(description);
    actionDiv.appendChild(button);
    chatContainer.appendChild(actionDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

async function executeAction(action) {
    try {
        await Excel.run(async (context) => {
            if (action.type === 'formula') {
                const range = context.workbook.getSelectedRange();
                range.formulas = [[action.formula]];
            } else if (action.type === 'value') {
                const range = context.workbook.getSelectedRange();
                range.values = [[action.value]];
            }
            
            await context.sync();
            addMessage('✅ Successfully inserted into Excel!', 'ai');
        });
    } catch (error) {
        console.error('Error executing action:', error);
        addMessage('❌ Failed to insert into Excel. Please try again.', 'error');
    }
}

function clearChat() {
    const chatContainer = document.getElementById('chatContainer');
    chatContainer.innerHTML = `
        <div class="welcome-message">
            <div class="welcome-icon">👋</div>
            <div class="welcome-content">
                <h4>Welcome to ExcelAI Assistant!</h4>
                <p>I can help you analyze data, create formulas, and provide insights. Here are some things you can ask:</p>
                <ul class="example-prompts">
                    <li>"What's the correlation between columns A and B?"</li>
                    <li>"Create a formula to calculate the growth rate"</li>
                    <li>"What insights can you find in this data?"</li>
                    <li>"Suggest the best chart type for this data"</li>
                </ul>
            </div>
        </div>
    `;
    chatHistory = [];
}

function showError(message) {
    addMessage(message, 'error');
}

// Initialize status on load
document.addEventListener('DOMContentLoaded', () => {
    updateStatus('Loading...', 'loading');
});
