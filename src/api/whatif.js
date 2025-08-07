const express = require('express');
const router = express.Router();
const WhatIfAnalysis = require('../features/whatIfAnalysis');

// What-If Scenario Analysis
router.post('/scenarios', async (req, res, next) => {
    try {
        const { selectedData, scenarioType = 'all' } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        let scenarios;
        
        switch (scenarioType) {
            case 'optimistic':
                scenarios = { optimistic: WhatIfAnalysis.createOptimisticScenario(selectedData) };
                break;
            case 'pessimistic':
                scenarios = { pessimistic: WhatIfAnalysis.createPessimisticScenario(selectedData) };
                break;
            case 'realistic':
                scenarios = { realistic: WhatIfAnalysis.createRealisticScenario(selectedData) };
                break;
            default:
                scenarios = WhatIfAnalysis.generateScenarios(selectedData);
        }

        res.json({
            success: true,
            scenarios,
            dataContext: {
                range: selectedData.address,
                size: `${selectedData.rowCount}×${selectedData.columnCount}`
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Sensitivity Analysis
router.post('/sensitivity', async (req, res, next) => {
    try {
        const { selectedData, variable, changeRange } = req.body;
        
        if (!selectedData || !variable) {
            return res.status(400).json({
                success: false,
                error: 'Selected data and variable are required'
            });
        }

        const analysis = WhatIfAnalysis.performSensitivityAnalysis(
            selectedData, 
            variable, 
            changeRange
        );

        res.json({
            success: true,
            sensitivityAnalysis: analysis,
            recommendations: analysis.recommendations,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Goal Seek Analysis
router.post('/goalseek', async (req, res, next) => {
    try {
        const { selectedData, targetValue, changingCell } = req.body;
        
        if (!selectedData || !targetValue || !changingCell) {
            return res.status(400).json({
                success: false,
                error: 'Selected data, target value, and changing cell are required'
            });
        }

        const goalSeekAnalysis = WhatIfAnalysis.generateGoalSeekAnalysis(
            selectedData,
            targetValue,
            changingCell
        );

        res.json({
            success: true,
            goalSeekAnalysis,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Data Table Setup
router.post('/datatable', async (req, res, next) => {
    try {
        const { selectedData, inputCells, resultCells } = req.body;
        
        if (!selectedData || !inputCells || !resultCells) {
            return res.status(400).json({
                success: false,
                error: 'Selected data, input cells, and result cells are required'
            });
        }

        const dataTable = WhatIfAnalysis.createDataTable(
            selectedData,
            inputCells,
            resultCells
        );

        res.json({
            success: true,
            dataTable,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Monte Carlo Simulation Setup
router.post('/montecarlo', async (req, res, next) => {
    try {
        const { selectedData, iterations } = req.body;
        
        if (!selectedData) {
            return res.status(400).json({
                success: false,
                error: 'Selected data is required'
            });
        }

        const monteCarloInputs = WhatIfAnalysis.generateMonteCarloInputs(
            selectedData,
            iterations
        );

        res.json({
            success: true,
            monteCarloSetup: monteCarloInputs,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

module.exports = router;
