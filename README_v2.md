# Excel AI Agent 2.0 🤖📊

> **Advanced AI-Powered Excel Add-in for Business Intelligence & Data Analysis**

An enterprise-grade Excel Add-in that transforms spreadsheet data into actionable business insights using artificial intelligence, advanced analytics, and sophisticated modeling capabilities.

## 🚀 Key Features

### Core Capabilities
- 🧠 **AI-Powered Analysis** - Natural language data analysis with GPT integration
- 📈 **What-If Analysis** - Comprehensive scenario modeling and sensitivity analysis
- 🔮 **Predictive Analytics** - Trend analysis and forecasting with confidence intervals
- 💡 **Business Intelligence** - Automated insights and pattern recognition
- 🧮 **Smart Formula Generation** - AI-assisted Excel formula creation and validation
- 📊 **Statistical Analysis** - Advanced statistical functions and data quality assessment

### Advanced Features
- **Scenario Planning** - Optimistic, realistic, and pessimistic scenario generation
- **Monte Carlo Simulation** - Probabilistic modeling with uncertainty analysis
- **Goal Seek Analysis** - Reverse engineering to find required inputs
- **Correlation Analysis** - Multi-variable relationship detection
- **Outlier Detection** - Automatic anomaly identification and handling
- **Trend Forecasting** - Time series analysis and future value prediction

## 🏗️ Architecture

### Backend Services
```
src/
├── server.js                 # Express server with advanced routing
├── features/
│   ├── aiService.js          # AI processing with intent recognition
│   ├── whatIfAnalysis.js     # Scenario modeling engine
│   ├── dataAnalyzer.js       # Statistical analysis suite
│   └── formulaGenerator.js   # Intelligent formula creation
└── api/
    ├── analyze.js            # Data analysis endpoints
    ├── whatif.js             # What-if analysis API
    ├── formulas.js           # Formula generation API
    ├── scenarios.js          # Scenario planning API
    ├── predictions.js        # Predictive analytics API
    └── insights.js           # Business intelligence API
```

### Frontend Interface
```
public/
├── taskpane.html            # Enhanced UI with advanced features
├── css/taskpane.css         # Modern styling with tabs and grids
└── js/taskpane.js           # Advanced JavaScript with API integration
```

## 🛠️ Installation & Setup

### Prerequisites
- Node.js 18+ and npm 9+
- Excel 2016+ (Windows/Mac) or Excel Online
- OpenAI API key (optional, for enhanced AI features)

### Quick Start
```bash
# Clone the repository
git clone https://github.com/ParthPatelCa/Excel-AI-Agent.git
cd Excel-AI-Agent

# Install dependencies
npm install

# Configure environment
cp .env.example .env
# Edit .env with your settings

# Start the server
npm start
```

### Excel Add-in Setup
1. Open Excel and go to **Insert > Office Add-ins**
2. Choose **Upload My Add-in** and select `manifest.xml`
3. The Excel AI Agent will appear in the **Home** tab

## 🎯 Usage Guide

### Quick Actions
The interface provides 8 primary quick actions:

1. **📈 Analyze** - Comprehensive data analysis
2. **📋 Summarize** - Statistical summaries
3. **🎯 What-If** - Scenario analysis
4. **🔮 Predict** - Trend forecasting
5. **💡 Insights** - Business intelligence
6. **🧮 Formulas** - Smart formula generation
7. **📊 Charts** - Visualization recommendations
8. **🗂️ Scenarios** - Scenario planning

### Advanced Features Tabs

#### Analysis Tab
- **📊 Statistical Analysis** - Comprehensive statistical metrics
- **🔗 Correlations** - Multi-variable relationship analysis
- **🎯 Outlier Detection** - Anomaly identification
- **📈 Trend Analysis** - Pattern and trend identification

#### Modeling Tab
- **🎯 Scenario Modeling** - What-if scenario generation
- **🔮 Predictions** - Future value forecasting
- **📉 Sensitivity Analysis** - Variable impact assessment
- **🎯 Goal Seek** - Reverse goal calculation

#### Automation Tab
- **📚 Formula Library** - Pre-built formula templates
- **🤖 Auto Insights** - Automated analysis
- **🧠 AI Formulas** - Natural language formula generation
- **🔍 Pattern Recognition** - Data pattern identification

### Example Interactions

**Natural Language Analysis:**
```
"Analyze sales data for trends and correlations"
→ Provides statistical analysis, trend identification, and correlation matrix
```

**What-If Scenarios:**
```
"Create scenarios with 5%, 10%, and 15% growth rates"
→ Generates optimistic, realistic, and pessimistic scenarios
```

**Predictive Analytics:**
```
"Predict next 12 months based on historical data"
→ Provides trend-based forecasts with confidence intervals
```

## 📊 API Documentation

### Core Endpoints

#### Data Analysis API
```http
POST /api/analyze
Content-Type: application/json

{
  "selectedData": {
    "address": "A1:C10",
    "values": [[...]],
    "rowCount": 10,
    "columnCount": 3
  },
  "analysisType": "comprehensive"
}
```

#### What-If Analysis API
```http
POST /api/whatif/scenarios
Content-Type: application/json

{
  "scenarioData": {
    "baseValue": 100000,
    "variables": ["growth_rate", "cost_reduction"],
    "ranges": {
      "growth_rate": [0.05, 0.15],
      "cost_reduction": [0.02, 0.10]
    }
  }
}
```

#### Formula Generation API
```http
POST /api/formulas/generate
Content-Type: application/json

{
  "description": "Calculate NPV with 10% discount rate",
  "dataContext": {
    "range": "B1:B10",
    "type": "financial"
  }
}
```

### Response Format
All API endpoints return structured responses:

```json
{
  "success": true,
  "data": { /* analysis results */ },
  "metadata": {
    "timestamp": "2024-12-19T10:30:00.000Z",
    "processingTime": "150ms",
    "dataPoints": 100
  },
  "recommendations": ["..."],
  "formulas": { /* Excel formulas */ },
  "visualization": { /* chart suggestions */ }
}
```

## 🧮 Formula Integration

### Automated Formula Generation
The system generates Excel-compatible formulas for various scenarios:

**Financial Analysis:**
```excel
=NPV(0.1, B2:B11)  # Net Present Value
=IRR(B2:B11)       # Internal Rate of Return
=PMT(0.05/12, 60, 100000)  # Loan Payment
```

**Statistical Analysis:**
```excel
=CORREL(A2:A11, B2:B11)  # Correlation
=FORECAST.LINEAR(13, B2:B11, A2:A11)  # Linear Forecast
=CONFIDENCE(0.05, STDEV(A2:A11), COUNT(A2:A11))  # Confidence Interval
```

**What-If Analysis:**
```excel
=IF(scenario="Optimistic", value*1.15, IF(scenario="Pessimistic", value*0.85, value))
```

## 🔍 Advanced Analytics

### Statistical Functions
- **Descriptive Statistics**: Mean, median, mode, standard deviation, quartiles
- **Correlation Analysis**: Pearson correlation with significance testing
- **Regression Analysis**: Linear and polynomial regression
- **Hypothesis Testing**: T-tests, ANOVA, confidence intervals

### Predictive Modeling
- **Trend Analysis**: Linear and exponential trend identification
- **Time Series Forecasting**: ARIMA, exponential smoothing
- **Seasonal Decomposition**: Trend, seasonal, and residual components
- **Confidence Intervals**: Prediction uncertainty quantification

### Business Intelligence
- **Pattern Recognition**: Automatic pattern detection in data
- **Anomaly Detection**: Statistical outlier identification
- **Performance Metrics**: KPI calculation and tracking
- **Risk Assessment**: Scenario-based risk analysis

## 🛡️ Security & Performance

### Security Features
- **Rate Limiting**: API request throttling
- **Input Validation**: Comprehensive data validation
- **CORS Protection**: Cross-origin request security
- **Content Security Policy**: XSS prevention
- **Helmet Integration**: Security headers

### Performance Optimization
- **Compression**: Response compression for faster loading
- **Caching**: Intelligent result caching
- **Streaming**: Large dataset streaming
- **Memory Management**: Efficient memory usage

## 🧪 Testing

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run linting
npm run lint

# Fix linting issues
npm run lint:fix
```

## 📈 Deployment

### Local Development
```bash
npm run dev  # Uses nodemon for auto-restart
```

### Production Deployment
```bash
# Build for production
npm run build

# Start production server
npm start
```

### Environment Variables
```env
PORT=3000
OPENAI_API_KEY=your_openai_key_here
RATE_LIMIT_WINDOW_MS=900000
RATE_LIMIT_MAX_REQUESTS=100
LOG_LEVEL=info
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Commit changes: `git commit -m 'Add amazing feature'`
4. Push to branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## 📋 Changelog

### Version 2.0.0 (Latest)
- ✨ **New**: Comprehensive what-if analysis engine
- ✨ **New**: Advanced statistical analysis suite
- ✨ **New**: Predictive analytics and forecasting
- ✨ **New**: Business intelligence insights
- ✨ **New**: Smart formula generation
- ✨ **New**: Scenario planning and optimization
- 🔧 **Enhanced**: Modern tabbed UI interface
- 🔧 **Enhanced**: Advanced API architecture
- 🔧 **Enhanced**: Improved error handling

### Version 1.0.0
- 🎉 Initial release with basic chat functionality
- 📊 Basic data analysis capabilities
- 🤖 OpenAI integration

## 📞 Support

- **Documentation**: [Advanced Features Guide](docs/ADVANCED_FEATURES.md)
- **Issues**: [GitHub Issues](https://github.com/ParthPatelCa/Excel-AI-Agent/issues)
- **Discussions**: [GitHub Discussions](https://github.com/ParthPatelCa/Excel-AI-Agent/discussions)
- **Email**: support@excellai.com

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Office.js team for the Excel Add-in framework
- OpenAI for providing advanced AI capabilities
- The open-source community for excellent libraries and tools

---

**Made with ❤️ for Excel power users and data analysts worldwide**

*Transform your spreadsheets into intelligent decision-making tools with Excel AI Agent 2.0*
