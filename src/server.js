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

// Basic chat endpoint (ready for OpenAI integration)
app.post('/api/chat', async (req, res) => {
    try {
        const { message, selectedData } = req.body;
        
        if (!message || !selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Message and selected data are required'
            });
        }

        // TODO: Integrate OpenAI API here
        const response = `I received your message: "${message}" about data range ${selectedData.address}. 
        
To enable full AI functionality:
1. Add your OpenAI API key to .env file
2. Integrate OpenAI chat completion API
3. Add advanced features like formula generation and data analysis

Data Summary:
- Range: ${selectedData.address}
- Size: ${selectedData.rowCount} rows × ${selectedData.columnCount} columns`;

        res.json({
            success: true,
            response: response,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        console.error('Chat endpoint error:', error);
        res.status(500).json({
            success: false,
            error: 'Internal server error'
        });
    }
});

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
