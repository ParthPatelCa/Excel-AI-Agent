const express = require('express');
const router = express.Router();
const DataAnalyzer = require('../features/dataAnalyzer');
const WhatIfAnalysis = require('../features/whatIfAnalysis');

// Generate Predictions
router.post('/generate', async (req, res, next) => {
    try {
        const { selectedData, predictionType = 'trend', timeHorizon = 12, confidence = 0.95 } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required for predictions'
            });
        }

        let predictions;
        
        switch (predictionType) {
            case 'trend':
                predictions = generateTrendPredictions(selectedData, timeHorizon, confidence);
                break;
            case 'seasonal':
                predictions = generateSeasonalPredictions(selectedData, timeHorizon, confidence);
                break;
            case 'regression':
                predictions = generateRegressionPredictions(selectedData, timeHorizon, confidence);
                break;
            case 'moving_average':
                predictions = generateMovingAveragePredictions(selectedData, timeHorizon, confidence);
                break;
            case 'exponential':
                predictions = generateExponentialPredictions(selectedData, timeHorizon, confidence);
                break;
            default:
                predictions = generateComprehensivePredictions(selectedData, timeHorizon, confidence);
        }

        res.json({
            success: true,
            predictions,
            metadata: {
                dataPoints: selectedData.values ? selectedData.values.length : 0,
                predictionHorizon: timeHorizon,
                confidenceLevel: confidence,
                method: predictionType
            },
            formulas: {
                trend: `=TREND(${selectedData.address},ROW(${selectedData.address})-ROW(${selectedData.address}[1])+1,ROW(${selectedData.address}[1])+${timeHorizon})`,
                forecast: `=FORECAST.LINEAR(future_x,${selectedData.address},time_range)`,
                exponentialSmoothing: `=FORECAST.ETS(future_date,${selectedData.address},date_range,1,1,1)`
            },
            visualization: {
                chartType: 'line',
                series: [
                    { name: 'Historical Data', data: selectedData.values },
                    { name: 'Predictions', data: predictions.values },
                    { name: 'Confidence Band', data: predictions.confidenceBand }
                ]
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Time Series Analysis
router.post('/timeseries', async (req, res, next) => {
    try {
        const { selectedData, components = ['trend', 'seasonal', 'residual'] } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Time series data is required'
            });
        }

        const timeSeriesAnalysis = {
            decomposition: analyzeTimeSeriesComponents(selectedData, components),
            stationarity: testStationarity(selectedData),
            autocorrelation: calculateAutocorrelation(selectedData),
            seasonality: detectSeasonality(selectedData),
            outliers: DataAnalyzer.detectOutliers(selectedData),
            changePoints: detectChangePoints(selectedData)
        };

        // Generate forecasting recommendations
        const forecastingMethod = recommendForecastingMethod(timeSeriesAnalysis);

        res.json({
            success: true,
            timeSeriesAnalysis,
            forecastingMethod,
            recommendations: generateTimeSeriesRecommendations(timeSeriesAnalysis),
            implementation: {
                excelFunctions: [
                    'FORECAST.LINEAR() - Linear trend forecasting',
                    'FORECAST.ETS() - Exponential smoothing with seasonality',
                    'TREND() - Linear trend calculation',
                    'GROWTH() - Exponential trend calculation'
                ],
                additionalTools: [
                    'Analysis ToolPak for moving averages',
                    'Power Query for data preparation',
                    'Excel Charts for visualization'
                ]
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Forecast Accuracy
router.post('/accuracy', async (req, res, next) => {
    try {
        const { actualData, predictedData, metrics = ['mae', 'mse', 'mape', 'rmse'] } = req.body;
        
        if (!actualData || !predictedData) {
            return res.status(400).json({
                success: false,
                error: 'Both actual and predicted data are required'
            });
        }

        const accuracy = calculateForecastAccuracy(actualData, predictedData, metrics);
        const evaluation = evaluateForecastQuality(accuracy);

        res.json({
            success: true,
            accuracy,
            evaluation,
            improvement: {
                suggestions: generateImprovementSuggestions(accuracy),
                techniques: [
                    'Increase historical data length',
                    'Consider external variables',
                    'Apply data smoothing techniques',
                    'Use ensemble methods'
                ]
            },
            benchmarking: {
                naive: calculateNaiveForecastAccuracy(actualData),
                seasonal: calculateSeasonalForecastAccuracy(actualData),
                comparison: compareWithBenchmarks(accuracy, actualData)
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Scenario-Based Predictions
router.post('/scenarios', async (req, res, next) => {
    try {
        const { baseData, scenarios, variables } = req.body;
        
        if (!baseData || !scenarios) {
            return res.status(400).json({
                success: false,
                error: 'Base data and scenarios are required'
            });
        }

        const scenarioPredictions = scenarios.map(scenario => {
            const adjustedData = applyScenarioToData(baseData, scenario, variables);
            return {
                scenarioName: scenario.name,
                scenario: scenario,
                predictions: generateComprehensivePredictions(adjustedData, 12, 0.95),
                impact: calculateScenarioImpact(baseData, adjustedData),
                probability: scenario.probability || 0.33
            };
        });

        // Calculate weighted predictions
        const weightedPredictions = calculateWeightedPredictions(scenarioPredictions);
        
        // Risk analysis
        const riskAnalysis = {
            variance: calculatePredictionVariance(scenarioPredictions),
            worstCase: findWorstCaseScenario(scenarioPredictions),
            bestCase: findBestCaseScenario(scenarioPredictions),
            confidenceIntervals: calculateScenarioConfidenceIntervals(scenarioPredictions)
        };

        res.json({
            success: true,
            scenarioPredictions,
            weightedPredictions,
            riskAnalysis,
            recommendations: [
                'Focus on most probable scenarios',
                'Prepare contingency plans for worst case',
                'Monitor leading indicators',
                'Update scenarios regularly'
            ],
            monitoring: {
                keyIndicators: identifyKeyIndicators(scenarios),
                alertThresholds: calculateAlertThresholds(riskAnalysis),
                reviewSchedule: 'Monthly scenario updates recommended'
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Machine Learning Predictions
router.post('/ml', async (req, res, next) => {
    try {
        const { trainingData, features, target, algorithm = 'auto' } = req.body;
        
        if (!trainingData || !target) {
            return res.status(400).json({
                success: false,
                error: 'Training data and target variable are required'
            });
        }

        // Note: This is a simplified ML implementation
        // In production, you would integrate with actual ML libraries
        const mlPredictions = {
            algorithm: algorithm === 'auto' ? selectBestAlgorithm(trainingData, target) : algorithm,
            model: trainModel(trainingData, features, target, algorithm),
            predictions: generateMLPredictions(trainingData, features, target),
            featureImportance: calculateFeatureImportance(trainingData, features, target),
            modelMetrics: calculateModelMetrics(trainingData, target)
        };

        res.json({
            success: true,
            mlPredictions,
            interpretation: {
                accuracy: mlPredictions.modelMetrics.accuracy,
                topFeatures: mlPredictions.featureImportance.slice(0, 5),
                modelComplexity: assessModelComplexity(mlPredictions.model)
            },
            limitations: [
                'Requires sufficient historical data',
                'Assumes pattern continuation',
                'May not capture external events',
                'Needs regular model updates'
            ],
            excelIntegration: {
                note: 'Excel has limited ML capabilities',
                alternatives: [
                    'Use Excel forecasting functions',
                    'Export data to Python/R for ML',
                    'Consider Power BI for advanced analytics',
                    'Use Azure ML for cloud-based solutions'
                ]
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Helper Functions
function generateTrendPredictions(data, horizon, confidence) {
    // Simplified trend analysis
    const values = data.values || [];
    const n = values.length;
    
    if (n < 2) {
        throw new Error('Insufficient data for trend prediction');
    }
    
    // Calculate linear trend
    const x = Array.from({length: n}, (_, i) => i + 1);
    const y = values;
    
    const slope = calculateSlope(x, y);
    const intercept = calculateIntercept(x, y, slope);
    
    const predictions = [];
    const confidenceBand = [];
    
    for (let i = 1; i <= horizon; i++) {
        const futureX = n + i;
        const prediction = slope * futureX + intercept;
        predictions.push(prediction);
        
        // Simple confidence band calculation
        const error = Math.sqrt(calculateMSE(x, y, slope, intercept));
        const margin = 1.96 * error; // 95% confidence
        confidenceBand.push({
            lower: prediction - margin,
            upper: prediction + margin
        });
    }
    
    return {
        values: predictions,
        confidenceBand,
        trend: {
            slope,
            intercept,
            direction: slope > 0 ? 'increasing' : 'decreasing',
            strength: Math.abs(slope)
        }
    };
}

function generateComprehensivePredictions(data, horizon, confidence) {
    const trendPred = generateTrendPredictions(data, horizon, confidence);
    
    return {
        method: 'comprehensive',
        trend: trendPred,
        accuracy: estimateAccuracy(data),
        reliability: assessReliability(data, horizon)
    };
}

function calculateSlope(x, y) {
    const n = x.length;
    const sumX = x.reduce((a, b) => a + b, 0);
    const sumY = y.reduce((a, b) => a + b, 0);
    const sumXY = x.reduce((sum, xi, i) => sum + xi * y[i], 0);
    const sumX2 = x.reduce((sum, xi) => sum + xi * xi, 0);
    
    return (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
}

function calculateIntercept(x, y, slope) {
    const meanX = x.reduce((a, b) => a + b, 0) / x.length;
    const meanY = y.reduce((a, b) => a + b, 0) / y.length;
    return meanY - slope * meanX;
}

function calculateMSE(x, y, slope, intercept) {
    const predictions = x.map(xi => slope * xi + intercept);
    const squaredErrors = y.map((yi, i) => Math.pow(yi - predictions[i], 2));
    return squaredErrors.reduce((a, b) => a + b, 0) / y.length;
}

function estimateAccuracy(data) {
    // Simple accuracy estimation based on data characteristics
    const values = data.values || [];
    if (values.length < 10) return 'Low';
    
    const coefficient = calculateVariationCoefficient(values);
    if (coefficient < 0.1) return 'High';
    if (coefficient < 0.3) return 'Medium';
    return 'Low';
}

function calculateVariationCoefficient(values) {
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
    const stdDev = Math.sqrt(variance);
    return stdDev / mean;
}

function assessReliability(data, horizon) {
    const values = data.values || [];
    const dataLength = values.length;
    
    if (horizon > dataLength) return 'Low';
    if (horizon > dataLength / 2) return 'Medium';
    return 'High';
}

// Additional helper functions would be implemented here...
function analyzeTimeSeriesComponents(data, components) { return {}; }
function testStationarity(data) { return {}; }
function calculateAutocorrelation(data) { return {}; }
function detectSeasonality(data) { return {}; }
function detectChangePoints(data) { return {}; }
function recommendForecastingMethod(analysis) { return {}; }
function generateTimeSeriesRecommendations(analysis) { return []; }
function calculateForecastAccuracy(actual, predicted, metrics) { return {}; }
function evaluateForecastQuality(accuracy) { return {}; }
function generateImprovementSuggestions(accuracy) { return []; }
function calculateNaiveForecastAccuracy(data) { return {}; }
function calculateSeasonalForecastAccuracy(data) { return {}; }
function compareWithBenchmarks(accuracy, data) { return {}; }
function applyScenarioToData(baseData, scenario, variables) { return baseData; }
function calculateScenarioImpact(baseData, adjustedData) { return {}; }
function calculateWeightedPredictions(scenarioPredictions) { return {}; }
function calculatePredictionVariance(predictions) { return 0; }
function findWorstCaseScenario(predictions) { return {}; }
function findBestCaseScenario(predictions) { return {}; }
function calculateScenarioConfidenceIntervals(predictions) { return {}; }
function identifyKeyIndicators(scenarios) { return []; }
function calculateAlertThresholds(riskAnalysis) { return {}; }
function selectBestAlgorithm(data, target) { return 'linear_regression'; }
function trainModel(data, features, target, algorithm) { return {}; }
function generateMLPredictions(data, features, target) { return {}; }
function calculateFeatureImportance(data, features, target) { return []; }
function calculateModelMetrics(data, target) { return {}; }
function assessModelComplexity(model) { return 'medium'; }

module.exports = router;
