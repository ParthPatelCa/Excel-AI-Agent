const axios = require('axios');
const DataAnalyzer = require('./dataAnalyzer');
const WhatIfAnalysis = require('./whatIfAnalysis');
const FormulaGenerator = require('./formulaGenerator');

class AIService {
    static async processMessage(message, selectedData, chatHistory = []) {
        if (!process.env.OPENAI_API_KEY) {
            throw new Error('OpenAI API key not configured');
        }

        // Analyze the intent
        const intent = this.analyzeIntent(message);
        const dataContext = this.buildDataContext(selectedData);
        
        let response;
        let action = null;

        try {
            switch (intent.type) {
                case 'whatif':
                    response = await this.handleWhatIfAnalysis(message, selectedData, intent);
                    break;
                case 'analysis':
                    response = await this.handleDataAnalysis(message, selectedData);
                    break;
                case 'formula':
                    const formulaResult = await this.handleFormulaGeneration(message, selectedData);
                    response = formulaResult.response;
                    action = formulaResult.action;
                    break;
                case 'prediction':
                    response = await this.handlePrediction(message, selectedData);
                    break;
                case 'scenario':
                    response = await this.handleScenarioAnalysis(message, selectedData);
                    break;
                default:
                    response = await this.handleGeneralQuery(message, selectedData, chatHistory);
            }

            return { response, action };

        } catch (error) {
            console.error('AI Service error:', error);
            throw new Error('Failed to process AI request');
        }
    }

    static analyzeIntent(message) {
        const lowerMessage = message.toLowerCase();
        
        const intents = {
            whatif: ['what if', 'what would happen', 'scenario', 'suppose', 'if i change'],
            analysis: ['analyze', 'insights', 'patterns', 'trends', 'correlation', 'statistics'],
            formula: ['formula', 'calculate', 'create formula', 'function', 'sum', 'average'],
            prediction: ['predict', 'forecast', 'future', 'projection', 'trend'],
            scenario: ['scenario', 'compare', 'best case', 'worst case', 'sensitivity']
        };

        for (const [type, keywords] of Object.entries(intents)) {
            if (keywords.some(keyword => lowerMessage.includes(keyword))) {
                return { 
                    type, 
                    confidence: this.calculateConfidence(lowerMessage, keywords),
                    keywords: keywords.filter(k => lowerMessage.includes(k))
                };
            }
        }

        return { type: 'general', confidence: 0.5, keywords: [] };
    }

    static calculateConfidence(message, keywords) {
        const matches = keywords.filter(keyword => message.includes(keyword)).length;
        return Math.min(matches / keywords.length + 0.3, 1.0);
    }

    static buildDataContext(selectedData) {
        const stats = DataAnalyzer.getBasicStats(selectedData);
        
        return `
Data Context:
- Range: ${selectedData.address}
- Dimensions: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns
- Data Types: ${stats.dataTypes.join(', ')}
- Has Headers: ${stats.hasHeaders}
- Numeric Columns: ${stats.numericColumns}
- Sample Data: ${JSON.stringify(selectedData.values.slice(0, 3))}
`;
    }

    static async handleWhatIfAnalysis(message, selectedData, intent) {
        const systemPrompt = `You are a What-If Analysis expert for Excel. Analyze the user's scenario question and provide:

1. Clear explanation of the scenario
2. Step-by-step Excel formulas to model it
3. Expected outcomes and insights
4. Additional scenarios to consider

Focus on practical Excel solutions using functions like:
- Goal Seek analysis
- Data Tables
- Scenario Manager
- Sensitivity analysis formulas

Be specific about which cells to modify and what formulas to use.`;

        const userPrompt = `${message}

Data Context: ${this.buildDataContext(selectedData)}

Please provide a comprehensive what-if analysis with specific Excel formulas and expected outcomes.`;

        return await this.callOpenAI(systemPrompt, userPrompt);
    }

    static async handleDataAnalysis(message, selectedData) {
        const analysis = DataAnalyzer.performDeepAnalysis(selectedData);
        
        const systemPrompt = `You are a data analysis expert. Based on the provided data analysis results, create insights and actionable recommendations.

Focus on:
- Key patterns and trends
- Statistical significance
- Business implications
- Recommended actions
- Excel formulas for further analysis`;

        const userPrompt = `${message}

Analysis Results: ${JSON.stringify(analysis)}
Data Context: ${this.buildDataContext(selectedData)}

Provide clear, actionable insights with specific Excel recommendations.`;

        return await this.callOpenAI(systemPrompt, userPrompt);
    }

    static async handleFormulaGeneration(message, selectedData) {
        const systemPrompt = `You are an Excel formula expert. Generate accurate, efficient Excel formulas based on user requests.

Guidelines:
- Use proper Excel syntax with = prefix
- Prefer built-in functions over complex nested formulas
- Include error handling where appropriate
- Explain what each formula does
- Suggest alternatives when relevant

Always format formulas clearly and ensure they're safe for Excel use.`;

        const userPrompt = `${message}

Data Context: ${this.buildDataContext(selectedData)}

Please provide the exact Excel formula(s) needed, with explanations.`;

        const response = await this.callOpenAI(systemPrompt, userPrompt);
        const action = FormulaGenerator.extractFormula(response);

        return { response, action };
    }

    static async handlePrediction(message, selectedData) {
        const systemPrompt = `You are a predictive analytics expert for Excel. Help users create forecasts and predictions using Excel's built-in functions.

Focus on:
- FORECAST functions
- TREND analysis
- Linear/exponential regression
- Moving averages
- Seasonal decomposition

Provide specific Excel formulas and methodology explanations.`;

        const userPrompt = `${message}

Data Context: ${this.buildDataContext(selectedData)}

Please provide prediction methodology and specific Excel formulas for forecasting.`;

        return await this.callOpenAI(systemPrompt, userPrompt);
    }

    static async handleScenarioAnalysis(message, selectedData) {
        const scenarios = WhatIfAnalysis.generateScenarios(selectedData);
        
        const systemPrompt = `You are a scenario planning expert. Help users understand different business scenarios and their implications.

Create comprehensive scenario analysis including:
- Best case scenarios
- Worst case scenarios
- Most likely scenarios
- Risk assessments
- Mitigation strategies`;

        const userPrompt = `${message}

Generated Scenarios: ${JSON.stringify(scenarios)}
Data Context: ${this.buildDataContext(selectedData)}

Provide detailed scenario analysis with actionable insights.`;

        return await this.callOpenAI(systemPrompt, userPrompt);
    }

    static async handleGeneralQuery(message, selectedData, chatHistory) {
        const conversationContext = chatHistory
            .slice(-3)
            .map(entry => `${entry.type}: ${entry.message}`)
            .join('\n');

        const systemPrompt = `You are ExcelAI Assistant, an expert Excel AI helper. You can:

🔮 What-If Analysis & Scenarios
📊 Advanced Data Analysis & Insights  
🧮 Smart Formula Generation
📈 Predictive Analytics & Forecasting
🎯 Goal Seek & Optimization
📋 Data Validation & Quality Assessment
🔗 Data Relationships & Correlations

Guidelines:
- Always provide actionable Excel solutions
- Use proper Excel function syntax
- Explain complex concepts simply
- Focus on practical business applications
- Suggest follow-up questions when helpful`;

        const userPrompt = `${message}

Data Context: ${this.buildDataContext(selectedData)}

Recent Conversation:
${conversationContext}

Please provide helpful, specific assistance with Excel formulas and data analysis.`;

        return await this.callOpenAI(systemPrompt, userPrompt);
    }

    static async callOpenAI(systemPrompt, userPrompt) {
        try {
            const response = await axios.post('https://api.openai.com/v1/chat/completions', {
                model: process.env.OPENAI_MODEL || 'gpt-3.5-turbo',
                messages: [
                    { role: 'system', content: systemPrompt },
                    { role: 'user', content: userPrompt }
                ],
                max_tokens: parseInt(process.env.OPENAI_MAX_TOKENS) || 1500,
                temperature: parseFloat(process.env.OPENAI_TEMPERATURE) || 0.7,
                presence_penalty: 0.1,
                frequency_penalty: 0.1
            }, {
                headers: {
                    'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
                    'Content-Type': 'application/json'
                },
                timeout: 30000
            });

            return response.data.choices[0].message.content;

        } catch (error) {
            console.error('OpenAI API error:', error.response?.data || error.message);
            throw new Error('Failed to get AI response');
        }
    }
}

module.exports = AIService;
