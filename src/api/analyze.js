const express = require('express');
const router = express.Router();
const DataAnalyzer = require('../features/dataAnalyzer');

// Comprehensive Data Analysis
router.post('/', async (req, res, next) => {
    try {
        const { selectedData, analysisType = 'comprehensive' } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        let analysis;
        
        switch (analysisType) {
            case 'basic':
                analysis = DataAnalyzer.getBasicStats(selectedData);
                break;
            case 'statistical':
                analysis = DataAnalyzer.calculateStatistics(selectedData);
                break;
            case 'quality':
                analysis = DataAnalyzer.assessDataQuality(selectedData);
                break;
            case 'patterns':
                analysis = DataAnalyzer.detectPatterns(selectedData);
                break;
            case 'correlations':
                analysis = DataAnalyzer.calculateCorrelations(selectedData);
                break;
            case 'outliers':
                analysis = DataAnalyzer.detectOutliers(selectedData);
                break;
            case 'trends':
                analysis = DataAnalyzer.analyzeTrends(selectedData);
                break;
            default:
                analysis = DataAnalyzer.performDeepAnalysis(selectedData);
        }

        res.json({
            success: true,
            analysis,
            dataContext: {
                range: selectedData.address,
                dimensions: `${selectedData.rowCount} rows × ${selectedData.columnCount} columns`,
                analysisType
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Statistical Summary
router.post('/statistics', async (req, res, next) => {
    try {
        const { selectedData } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const statistics = DataAnalyzer.calculateStatistics(selectedData);

        res.json({
            success: true,
            statistics,
            summary: {
                range: selectedData.address,
                numericColumns: Object.keys(statistics).length,
                insights: DataAnalyzer.generateRecommendations(selectedData)
                    .filter(r => r.category === 'Statistical Analysis')
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Data Quality Assessment
router.post('/quality', async (req, res, next) => {
    try {
        const { selectedData } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const qualityAssessment = DataAnalyzer.assessDataQuality(selectedData);

        res.json({
            success: true,
            qualityAssessment,
            recommendations: qualityAssessment.overall.recommendations,
            actionItems: [
                {
                    priority: 'high',
                    action: 'Address missing values',
                    formula: '=IFERROR(original_formula, "N/A")'
                },
                {
                    priority: 'medium',
                    action: 'Remove duplicates',
                    method: 'Data > Remove Duplicates'
                },
                {
                    priority: 'low',
                    action: 'Standardize formats',
                    tools: 'Data validation, Text to Columns'
                }
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Correlation Analysis
router.post('/correlations', async (req, res, next) => {
    try {
        const { selectedData, threshold = 0.5 } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const correlations = DataAnalyzer.calculateCorrelations(selectedData);
        
        // Filter by threshold
        const significantCorrelations = Object.entries(correlations)
            .filter(([key, corr]) => Math.abs(corr.coefficient) >= threshold)
            .reduce((obj, [key, corr]) => {
                obj[key] = corr;
                return obj;
            }, {});

        res.json({
            success: true,
            correlations: significantCorrelations,
            allCorrelations: correlations,
            insights: Object.entries(significantCorrelations).map(([key, corr]) => ({
                relationship: key,
                strength: corr.strength,
                interpretation: corr.interpretation,
                coefficient: corr.coefficient
            })),
            recommendations: [
                'Create scatter plots for strong correlations',
                'Investigate causal relationships',
                'Consider regression analysis for prediction'
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Outlier Detection
router.post('/outliers', async (req, res, next) => {
    try {
        const { selectedData, method = 'iqr' } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const outliers = DataAnalyzer.detectOutliers(selectedData);
        
        // Summarize outlier information
        const summary = Object.entries(outliers).map(([column, data]) => ({
            column,
            mildOutliers: data.mild.count,
            extremeOutliers: data.extreme.count,
            percentage: data.percentage,
            impact: data.impact,
            recommendations: this.getOutlierRecommendations(data)
        }));

        res.json({
            success: true,
            outliers,
            summary,
            formulas: {
                detectOutliers: '=IF(OR(A1<QUARTILE($A$1:$A$100,1)-1.5*IQR,A1>QUARTILE($A$1:$A$100,3)+1.5*IQR),"Outlier","Normal")',
                iqr: '=QUARTILE(range,3)-QUARTILE(range,1)',
                zScore: '=(A1-AVERAGE($A$1:$A$100))/STDEV($A$1:$A$100)'
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Trend Analysis
router.post('/trends', async (req, res, next) => {
    try {
        const { selectedData } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const trends = DataAnalyzer.analyzeTrends(selectedData);
        
        // Generate trend insights
        const insights = Object.entries(trends).map(([column, trend]) => ({
            column,
            direction: trend.linear.direction,
            strength: trend.linear.strength,
            description: trend.linear.description,
            volatility: trend.volatility,
            momentum: trend.momentum,
            formulas: {
                trendLine: `=TREND(${column}_range,ROW(${column}_range)-ROW(${column}_range[1])+1)`,
                slope: `=SLOPE(${column}_range,ROW(${column}_range)-ROW(${column}_range[1])+1)`,
                rSquared: `=RSQ(${column}_range,ROW(${column}_range)-ROW(${column}_range[1])+1)`
            }
        }));

        res.json({
            success: true,
            trends,
            insights,
            recommendations: [
                'Create trend line charts for visualization',
                'Use FORECAST function for predictions',
                'Monitor trend sustainability',
                'Consider seasonal adjustments'
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Helper method for outlier recommendations
function getOutlierRecommendations(outlierData) {
    const recommendations = [];
    
    if (outlierData.percentage > 10) {
        recommendations.push('High outlier rate - investigate data collection process');
    }
    
    if (outlierData.extreme.count > 0) {
        recommendations.push('Extreme outliers detected - verify data accuracy');
    }
    
    if (outlierData.impact === 'high') {
        recommendations.push('Consider robust statistical methods');
    }
    
    recommendations.push('Use conditional formatting to highlight outliers');
    
    return recommendations;
}

module.exports = router;
