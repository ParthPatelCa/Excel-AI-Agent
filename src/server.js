const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const morgan = require('morgan');
const rateLimit = require('express-rate-limit');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Security middleware
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            scriptSrc: ["'self'", "'unsafe-inline'", "https://appsforoffice.microsoft.com"],
            styleSrc: ["'self'", "'unsafe-inline'"],
            imgSrc: ["'self'", "data:", "https:"],
            connectSrc: ["'self'", "https://api.openai.com"],
        }
    },
    crossOriginEmbedderPolicy: false
}));

// Performance middleware
app.use(compression());
app.use(morgan('combined'));

// Rate limiting
const limiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000,
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100,
    message: {
        error: 'Too many requests from this IP, please try again later.'
    }
});

app.use('/api/', limiter);

// CORS configuration for Office Add-ins
app.use(cors({
    origin: [
        'https://excel.office.com',
        'https://excel.office365.com',
        'https://outlook.office.com',
        'https://excel-online.office.com'
    ],
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

// Body parsing middleware
app.use(express.json({ limit: process.env.MAX_REQUEST_SIZE || '10mb' }));
app.use(express.urlencoded({ extended: true, limit: process.env.MAX_REQUEST_SIZE || '10mb' }));

// Static files
app.use(express.static(path.join(__dirname, '../public')));

// Advanced API Routes
app.use('/api/analyze', require('./api/analyze'));
// Mount API routes
app.use('/api/whatif', require('./api/whatif'));
app.use('/api/analyze', require('./api/analyze'));
app.use('/api/formulas', require('./api/formulas'));
app.use('/api/scenarios', require('./api/scenarios'));
app.use('/api/predictions', require('./api/predictions'));
app.use('/api/insights', require('./api/insights'));
app.use('/api/formulas', require('./api/formulas'));
app.use('/api/scenarios', require('./api/scenarios'));
app.use('/api/predictions', require('./api/predictions'));
app.use('/api/insights', require('./api/insights'));

// Import advanced features
const DataAnalyzer = require('./features/dataAnalyzer');
const WhatIfAnalysis = require('./features/whatIfAnalysis');
const FormulaGenerator = require('./features/formulaGenerator');
const AIService = require('./features/aiService');

// Enhanced chat endpoint with AI integration
app.post('/api/chat', async (req, res) => {
    try {
        const { message, selectedData, chatHistory = [] } = req.body;
        
        if (!message || !selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Message and selected data are required'
            });
        }

        // Use AI service if available, otherwise fallback
        let response;
        let action = null;
        
        if (process.env.OPENAI_API_KEY) {
            const aiResult = await AIService.processMessage(message, selectedData, chatHistory);
            response = aiResult.response;
            action = aiResult.action;
        } else {
            response = await generateFallbackResponse(message, selectedData);
        }

        res.json({
            success: true,
            response: response,
            action: action,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        console.error('Chat endpoint error:', error);
        res.status(500).json({
            success: false,
            error: error.message || 'Internal server error'
        });
    }
});

// Fallback response when OpenAI is not configured
async function generateFallbackResponse(message, selectedData) {
    const lowerMessage = message.toLowerCase();
    
    if (lowerMessage.includes('what if') || lowerMessage.includes('scenario')) {
        return `What-If Analysis for ${selectedData.address}:
        
I can help you perform scenario analysis! Here are some examples:
• "What if I increase column A values by 10%?"
• "What if column B costs double?"
• "What would happen if sales grew by 15%?"

To enable full AI-powered what-if analysis, add your OpenAI API key to the .env file.

Current data: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns`;
    }
    
    if (lowerMessage.includes('average') || lowerMessage.includes('mean')) {
        return `For range ${selectedData.address}, I can help calculate averages.
        
Try these formulas:
• =AVERAGE(${selectedData.address}) - Simple average
• =AVERAGEIF(${selectedData.address},">0") - Average of positive values only

Data Summary: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns`;
    }
    
    if (lowerMessage.includes('sum') || lowerMessage.includes('total')) {
        return `For range ${selectedData.address}, here are sum formulas:
        
• =SUM(${selectedData.address}) - Sum all values
• =SUMIF(${selectedData.address},">0") - Sum positive values only
• =SUMPRODUCT(${selectedData.address}) - Product sum

Current selection: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns`;
    }
    
    return `I received your message about "${message}" for range ${selectedData.address}.

To unlock full AI capabilities including:
🔮 What-If Analysis & Scenarios
📊 Advanced Data Insights  
🧮 Smart Formula Generation
📈 Predictive Analytics
🎯 Goal Seek Analysis

Please add your OpenAI API key to the .env file.

Current data: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns`;
}

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({
        status: 'healthy',
        timestamp: new Date().toISOString(),
        version: '2.0.0',
        environment: process.env.NODE_ENV || 'development',
        features: {
            openai: !!process.env.OPENAI_API_KEY,
            moderation: process.env.ENABLE_MODERATION === 'true'
        }
    });
});

// API documentation endpoint
app.get('/api', (req, res) => {
    res.json({
        name: 'Excel AI Agent API',
        version: '2.0.0',
        endpoints: {
            chat: 'POST /api/chat - AI chat interface',
            health: 'GET /health - Health check'
        },
        documentation: 'https://github.com/ParthPatelCa/Excel-AI-Agent#readme'
    });
});

// Serve main taskpane
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../public/taskpane.html'));
});

// 404 handler
app.use('*', (req, res) => {
    res.status(404).json({
        error: 'Endpoint not found',
        path: req.originalUrl,
        method: req.method
    });
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    process.exit(0);
});

process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    process.exit(0);
});

// Start server
const server = app.listen(PORT, () => {
    console.log(`🚀 Excel AI Agent server running on port ${PORT}`);
    console.log(`📊 Task pane available at: http://localhost:${PORT}`);
    console.log(`🔧 Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log(`📝 Health check: http://localhost:${PORT}/health`);
    console.log(`🤖 OpenAI configured: ${!!process.env.OPENAI_API_KEY}`);
});

module.exports = app;
