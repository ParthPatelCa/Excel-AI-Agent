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

    // Advanced feature buttons
    document.querySelectorAll('.feature-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const action = e.currentTarget.dataset.action;
            handleQuickAction(action);
        });
    });

    // Tab functionality
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const tabName = e.currentTarget.dataset.tab;
            switchTab(tabName);
        });
    });

    // Initialize
    refreshSelectedData();
    updateInputState();
}

function switchTab(tabName) {
    // Remove active class from all tabs and panels
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));
    
    // Add active class to selected tab and panel
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(`${tabName}-tab`).classList.add('active');
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

    // Handle advanced feature actions
    if (action.startsWith('api:')) {
        await handleAdvancedAction(action);
        return;
    }

    const prompts = {
        analyze: 'Perform a comprehensive analysis of this data including statistical insights, trends, outliers, and correlations.',
        summarize: 'Generate a detailed summary of this data including descriptive statistics, data quality assessment, and key findings.',
        chart: 'What type of chart would be best for visualizing this data? Explain why and provide specific recommendations.',
        formulas: 'Suggest useful Excel formulas I could use with this data, including statistical, financial, and analytical formulas.',
        whatif: 'Help me set up what-if analysis scenarios for this data. Generate optimistic, realistic, and pessimistic scenarios.',
        predict: 'Analyze trends in this data and provide predictions for future values with confidence intervals.',
        insights: 'Generate business insights from this data including patterns, opportunities, and recommendations.',
        scenarios: 'Create scenario planning templates for this data with different business assumptions.'
    };

    const message = prompts[action];
    if (message) {
        document.getElementById('chatInput').value = message;
        updateInputState();
        await handleSendMessage();
    }
}

async function handleAdvancedAction(action) {
    const [, endpoint, type] = action.split(':');
    showLoadingState(true);
    
    try {
        let apiUrl, requestBody;
        
        switch (endpoint) {
            case 'analyze':
                apiUrl = '/api/analyze';
                requestBody = { 
                    selectedData: selectedRange, 
                    analysisType: type || 'comprehensive' 
                };
                break;
                
            case 'whatif':
                apiUrl = '/api/whatif/scenarios';
                requestBody = { 
                    scenarioData: selectedRange 
                };
                break;
                
            case 'predictions':
                apiUrl = '/api/predictions/generate';
                requestBody = { 
                    selectedData: selectedRange, 
                    predictionType: type || 'trend',
                    timeHorizon: 12
                };
                break;
                
            case 'insights':
                apiUrl = '/api/insights/generate';
                requestBody = { 
                    selectedData: selectedRange, 
                    insightType: type || 'comprehensive' 
                };
                break;
                
            case 'formulas':
                if (type) {
                    apiUrl = `/api/formulas/${type}`;
                    requestBody = { 
                        parameters: {
                            range: selectedRange.address,
                            data: selectedRange.values
                        }
                    };
                } else {
                    apiUrl = '/api/formulas/library';
                    await displayFormulaLibrary();
                    return;
                }
                break;
                
            default:
                throw new Error(`Unknown endpoint: ${endpoint}`);
        }
        
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(requestBody)
        });
        
        const data = await response.json();
        
        if (data.success) {
            displayAdvancedResults(data, endpoint);
        } else {
            addMessage(`Error: ${data.error}`, 'error');
        }
        
    } catch (error) {
        console.error('Advanced action error:', error);
        addMessage(`Failed to execute ${endpoint} analysis: ${error.message}`, 'error');
    } finally {
        showLoadingState(false);
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

function displayAdvancedResults(data, endpoint) {
    const chatContainer = document.getElementById('chatContainer');
    const resultDiv = document.createElement('div');
    resultDiv.className = 'message ai-message advanced-results';
    
    let content = '';
    
    switch (endpoint) {
        case 'analyze':
            content = createAnalysisDisplay(data);
            break;
        case 'whatif':
            content = createWhatIfDisplay(data);
            break;
        case 'predictions':
            content = createPredictionsDisplay(data);
            break;
        case 'insights':
            content = createInsightsDisplay(data);
            break;
        case 'formulas':
            content = createFormulaDisplay(data);
            break;
        default:
            content = `<pre>${JSON.stringify(data, null, 2)}</pre>`;
    }
    
    resultDiv.innerHTML = content;
    chatContainer.appendChild(resultDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

function createAnalysisDisplay(data) {
    const analysis = data.analysis || {};
    const statistics = analysis.statistics || {};
    
    return `
        <div class="analysis-results">
            <h4>📊 Data Analysis Results</h4>
            
            ${statistics.basicStats ? `
                <div class="stats-section">
                    <h5>Basic Statistics</h5>
                    <div class="stats-grid">
                        ${Object.entries(statistics.basicStats).map(([column, stats]) => `
                            <div class="stat-item">
                                <strong>${column}:</strong>
                                <ul>
                                    <li>Mean: ${stats.mean?.toFixed(2) || 'N/A'}</li>
                                    <li>Median: ${stats.median?.toFixed(2) || 'N/A'}</li>
                                    <li>Std Dev: ${stats.standardDeviation?.toFixed(2) || 'N/A'}</li>
                                </ul>
                            </div>
                        `).join('')}
                    </div>
                </div>
            ` : ''}
            
            ${analysis.insights ? `
                <div class="insights-section">
                    <h5>Key Insights</h5>
                    <ul>
                        ${analysis.insights.map(insight => `<li>${insight}</li>`).join('')}
                    </ul>
                </div>
            ` : ''}
            
            ${analysis.recommendations ? `
                <div class="recommendations-section">
                    <h5>Recommendations</h5>
                    <ul>
                        ${analysis.recommendations.map(rec => `<li>${rec}</li>`).join('')}
                    </ul>
                </div>
            ` : ''}
        </div>
    `;
}

function createWhatIfDisplay(data) {
    const scenarios = data.scenarios || [];
    
    return `
        <div class="whatif-results">
            <h4>🎯 What-If Analysis</h4>
            
            <div class="scenarios-grid">
                ${scenarios.map(scenario => `
                    <div class="scenario-card">
                        <h5>${scenario.name}</h5>
                        <p><strong>Outcome:</strong> ${scenario.expectedOutcome}</p>
                        <p><strong>Probability:</strong> ${(scenario.probability * 100).toFixed(1)}%</p>
                        ${scenario.assumptions ? `
                            <details>
                                <summary>Assumptions</summary>
                                <ul>
                                    ${scenario.assumptions.map(assumption => `<li>${assumption}</li>`).join('')}
                                </ul>
                            </details>
                        ` : ''}
                    </div>
                `).join('')}
            </div>
            
            ${data.recommendations ? `
                <div class="recommendations-section">
                    <h5>Implementation Steps</h5>
                    <ol>
                        ${data.recommendations.map(rec => `<li>${rec}</li>`).join('')}
                    </ol>
                </div>
            ` : ''}
        </div>
    `;
}

function createPredictionsDisplay(data) {
    const predictions = data.predictions || {};
    
    return `
        <div class="predictions-results">
            <h4>🔮 Predictions</h4>
            
            <div class="prediction-summary">
                <p><strong>Method:</strong> ${predictions.method || 'Trend Analysis'}</p>
                <p><strong>Accuracy:</strong> ${predictions.accuracy || 'Medium'}</p>
                <p><strong>Reliability:</strong> ${predictions.reliability || 'Medium'}</p>
            </div>
            
            ${predictions.values ? `
                <div class="prediction-values">
                    <h5>Predicted Values</h5>
                    <div class="values-grid">
                        ${predictions.values.slice(0, 6).map((value, index) => `
                            <div class="value-item">
                                <strong>Period ${index + 1}:</strong> ${value.toFixed(2)}
                            </div>
                        `).join('')}
                    </div>
                </div>
            ` : ''}
            
            ${predictions.trend ? `
                <div class="trend-info">
                    <h5>Trend Analysis</h5>
                    <p><strong>Direction:</strong> ${predictions.trend.direction}</p>
                    <p><strong>Slope:</strong> ${predictions.trend.slope?.toFixed(4) || 'N/A'}</p>
                </div>
            ` : ''}
            
            <div class="formula-section">
                <h5>Excel Formulas</h5>
                <div class="formula-item">
                    <code>=FORECAST.LINEAR(future_x, known_y, known_x)</code>
                    <button onclick="copyToClipboard('=FORECAST.LINEAR(future_x, known_y, known_x)')">Copy</button>
                </div>
            </div>
        </div>
    `;
}

function createInsightsDisplay(data) {
    const insights = data.insights || {};
    const keyFindings = data.keyFindings || [];
    
    return `
        <div class="insights-results">
            <h4>💡 Business Insights</h4>
            
            ${keyFindings.length > 0 ? `
                <div class="key-findings">
                    <h5>Key Findings</h5>
                    ${keyFindings.map(finding => `
                        <div class="finding-item">
                            <strong>${finding.category}:</strong> ${finding.finding}
                            <span class="confidence">Confidence: ${(finding.confidence * 100).toFixed(0)}%</span>
                        </div>
                    `).join('')}
                </div>
            ` : ''}
            
            ${data.recommendations ? `
                <div class="recommendations-section">
                    <h5>Actionable Recommendations</h5>
                    ${data.recommendations.slice(0, 5).map(rec => `
                        <div class="recommendation-item">
                            <div class="rec-header">
                                <strong>${rec.action}</strong>
                                <span class="priority priority-${rec.priority}">${rec.priority}</span>
                            </div>
                            <p>${rec.insight}</p>
                            <small>Impact: ${rec.impact} | Effort: ${rec.effort} | Timeline: ${rec.timeline}</small>
                        </div>
                    `).join('')}
                </div>
            ` : ''}
            
            ${data.visualizations ? `
                <div class="visualization-suggestions">
                    <h5>Suggested Visualizations</h5>
                    <div class="viz-grid">
                        ${data.visualizations.map(viz => `
                            <div class="viz-item">
                                <strong>${viz.type}:</strong> ${viz.description}
                            </div>
                        `).join('')}
                    </div>
                </div>
            ` : ''}
        </div>
    `;
}

function createFormulaDisplay(data) {
    const formula = data.formula || {};
    
    return `
        <div class="formula-results">
            <h4>📝 Formula Generator</h4>
            
            <div class="formula-main">
                <h5>Generated Formula</h5>
                <div class="formula-code">
                    <code>${formula.formula || 'No formula generated'}</code>
                    <button onclick="copyToClipboard('${formula.formula}')">Copy</button>
                </div>
            </div>
            
            ${formula.description ? `
                <div class="formula-explanation">
                    <h5>Explanation</h5>
                    <p>${formula.description}</p>
                </div>
            ` : ''}
            
            ${formula.examples ? `
                <div class="formula-examples">
                    <h5>Examples</h5>
                    <ul>
                        ${formula.examples.map(example => `<li>${example}</li>`).join('')}
                    </ul>
                </div>
            ` : ''}
            
            ${data.validation ? `
                <div class="formula-validation">
                    <h5>Validation</h5>
                    <p class="${data.validation.isValid ? 'valid' : 'invalid'}">
                        ${data.validation.isValid ? '✅ Formula is valid' : '❌ Formula has issues'}
                    </p>
                    ${data.validation.errors ? `
                        <ul class="errors">
                            ${data.validation.errors.map(error => `<li>${error}</li>`).join('')}
                        </ul>
                    ` : ''}
                </div>
            ` : ''}
        </div>
    `;
}

async function displayFormulaLibrary() {
    try {
        const response = await fetch('/api/formulas/library');
        const data = await response.json();
        
        if (data.success) {
            const content = `
                <div class="formula-library">
                    <h4>📚 Formula Library</h4>
                    
                    ${Object.entries(data.library).map(([category, formulas]) => `
                        <div class="formula-category">
                            <h5>${category.charAt(0).toUpperCase() + category.slice(1)} Formulas</h5>
                            <div class="formulas-grid">
                                ${formulas.map(formula => `
                                    <div class="formula-item">
                                        <strong>${formula.name}</strong>
                                        <p>${formula.description}</p>
                                        <code>${formula.syntax}</code>
                                        <button onclick="insertFormula('${formula.syntax}')">Insert</button>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
            
            addMessage(content, 'ai');
        }
    } catch (error) {
        addMessage('Failed to load formula library', 'error');
    }
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        // Show temporary success message
        const msg = document.createElement('div');
        msg.textContent = 'Copied to clipboard!';
        msg.style.cssText = 'position:fixed;top:20px;right:20px;background:#4caf50;color:white;padding:10px;border-radius:4px;z-index:1000;';
        document.body.appendChild(msg);
        setTimeout(() => msg.remove(), 2000);
    });
}

async function insertFormula(formula) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.formulas = [[formula]];
            await context.sync();
            addMessage('✅ Formula inserted successfully!', 'ai');
        });
    } catch (error) {
        addMessage('❌ Failed to insert formula', 'error');
    }
}

// Initialize status on load
document.addEventListener('DOMContentLoaded', () => {
    updateStatus('Loading...', 'loading');
});
