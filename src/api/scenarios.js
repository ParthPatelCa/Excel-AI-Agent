const express = require('express');
const router = express.Router();
const WhatIfAnalysis = require('../features/whatIfAnalysis');

// Create Scenario
router.post('/create', async (req, res, next) => {
    try {
        const { scenarioData, scenarioType = 'single' } = req.body;
        
        if (!scenarioData) {
            return res.status(400).json({
                success: false,
                error: 'Scenario data is required'
            });
        }

        let scenarios;
        
        switch (scenarioType) {
            case 'single':
                scenarios = [scenarioData];
                break;
            case 'multiple':
                scenarios = WhatIfAnalysis.generateScenarios(scenarioData);
                break;
            case 'optimistic':
                scenarios = WhatIfAnalysis.generateOptimisticScenario(scenarioData);
                break;
            case 'pessimistic':
                scenarios = WhatIfAnalysis.generatePessimisticScenario(scenarioData);
                break;
            case 'realistic':
                scenarios = WhatIfAnalysis.generateRealisticScenario(scenarioData);
                break;
            default:
                scenarios = WhatIfAnalysis.generateScenarios(scenarioData);
        }

        // Generate scenario comparison
        const comparison = WhatIfAnalysis.compareScenarios(scenarios);

        res.json({
            success: true,
            scenarios,
            comparison,
            recommendations: [
                'Use scenario manager to save scenarios',
                'Create charts to visualize scenario impacts',
                'Document assumptions for each scenario',
                'Regularly update scenarios with new data'
            ],
            implementation: {
                excelSteps: [
                    'Go to Data > What-If Analysis > Scenario Manager',
                    'Click Add to create new scenario',
                    'Define changing cells and values',
                    'Add scenario summary report'
                ],
                formulas: {
                    scenarioCell: '=IF(scenario_flag="Optimistic",optimistic_value,IF(scenario_flag="Pessimistic",pessimistic_value,realistic_value))',
                    dynamicScenario: '=CHOOSE(scenario_number,scenario1_value,scenario2_value,scenario3_value)'
                }
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Compare Scenarios
router.post('/compare', async (req, res, next) => {
    try {
        const { scenarios, comparisonMetrics } = req.body;
        
        if (!scenarios || scenarios.length < 2) {
            return res.status(400).json({
                success: false,
                error: 'At least two scenarios are required for comparison'
            });
        }

        const comparison = WhatIfAnalysis.compareScenarios(scenarios);
        const analysis = WhatIfAnalysis.analyzeScenarioImpact(scenarios, comparisonMetrics);

        // Generate insights
        const insights = {
            bestCase: scenarios.reduce((best, current) => 
                current.expectedOutcome > best.expectedOutcome ? current : best
            ),
            worstCase: scenarios.reduce((worst, current) => 
                current.expectedOutcome < worst.expectedOutcome ? current : worst
            ),
            mostLikely: scenarios.find(s => s.probability === Math.max(...scenarios.map(sc => sc.probability))),
            riskAssessment: {
                variance: analysis.variance,
                standardDeviation: Math.sqrt(analysis.variance),
                confidenceInterval: analysis.confidenceInterval
            }
        };

        res.json({
            success: true,
            comparison,
            analysis,
            insights,
            visualization: {
                chartType: 'tornado',
                data: scenarios.map(s => ({
                    name: s.name,
                    value: s.expectedOutcome,
                    probability: s.probability
                }))
            },
            recommendations: [
                `Focus on ${insights.bestCase.name} for maximum returns`,
                `Prepare contingency plans for ${insights.worstCase.name}`,
                `Most likely outcome: ${insights.mostLikely.name}`,
                'Consider portfolio approach to manage risk'
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Scenario Planning
router.post('/planning', async (req, res, next) => {
    try {
        const { baselineData, variables, constraints = {}, timeHorizon = 12 } = req.body;
        
        if (!baselineData || !variables) {
            return res.status(400).json({
                success: false,
                error: 'Baseline data and variables are required'
            });
        }

        // Generate planning scenarios
        const planningScenarios = {
            conservative: WhatIfAnalysis.generateConservativeScenario(baselineData, variables, constraints),
            moderate: WhatIfAnalysis.generateModerateScenario(baselineData, variables, constraints),
            aggressive: WhatIfAnalysis.generateAggressiveScenario(baselineData, variables, constraints),
            stressTest: WhatIfAnalysis.generateStressTestScenario(baselineData, variables, constraints)
        };

        // Project scenarios over time horizon
        const projections = Object.entries(planningScenarios).map(([name, scenario]) => ({
            name,
            scenario,
            projection: WhatIfAnalysis.projectScenario(scenario, timeHorizon)
        }));

        // Strategic recommendations
        const strategy = {
            shortTerm: projections.filter(p => p.projection.timeline <= 3),
            mediumTerm: projections.filter(p => p.projection.timeline <= 12),
            longTerm: projections.filter(p => p.projection.timeline > 12),
            keyMilestones: WhatIfAnalysis.identifyMilestones(projections),
            riskMitigation: WhatIfAnalysis.generateRiskMitigation(planningScenarios.stressTest)
        };

        res.json({
            success: true,
            planningScenarios,
            projections,
            strategy,
            implementation: {
                excelTemplate: {
                    sheets: ['Baseline', 'Variables', 'Scenarios', 'Projections', 'Dashboard'],
                    keyFormulas: [
                        '=Baseline_Value * (1 + Variable_Change)',
                        '=IF(Time_Period<=Short_Term,Conservative_Rate,Aggressive_Rate)',
                        '=NPV(Discount_Rate,Projection_Values)'
                    ]
                },
                monitoringPlan: [
                    'Monthly review of key variables',
                    'Quarterly scenario updates',
                    'Annual strategic planning review'
                ]
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Scenario Templates
router.get('/templates', async (req, res, next) => {
    try {
        const { category = 'all', industry } = req.query;
        
        const templates = {
            financial: [
                {
                    name: 'Revenue Growth Scenarios',
                    description: 'Model different revenue growth rates',
                    variables: ['growth_rate', 'market_size', 'market_share'],
                    scenarios: ['Conservative (5%)', 'Moderate (10%)', 'Aggressive (15%)']
                },
                {
                    name: 'Cost Reduction Scenarios',
                    description: 'Analyze impact of cost-saving initiatives',
                    variables: ['labor_costs', 'material_costs', 'overhead'],
                    scenarios: ['5% reduction', '10% reduction', '15% reduction']
                },
                {
                    name: 'Investment Returns',
                    description: 'Model different investment return rates',
                    variables: ['initial_investment', 'annual_return', 'time_horizon'],
                    scenarios: ['Bear Market', 'Normal Market', 'Bull Market']
                }
            ],
            operational: [
                {
                    name: 'Capacity Planning',
                    description: 'Plan for different demand levels',
                    variables: ['demand', 'capacity', 'utilization_rate'],
                    scenarios: ['Low Demand', 'Expected Demand', 'High Demand']
                },
                {
                    name: 'Supply Chain Disruption',
                    description: 'Model supply chain risk scenarios',
                    variables: ['supplier_availability', 'shipping_costs', 'inventory_levels'],
                    scenarios: ['Minor Disruption', 'Major Disruption', 'Crisis']
                }
            ],
            market: [
                {
                    name: 'Market Penetration',
                    description: 'Model market entry scenarios',
                    variables: ['market_size', 'penetration_rate', 'competition'],
                    scenarios: ['Slow Growth', 'Steady Growth', 'Rapid Growth']
                },
                {
                    name: 'Competitive Response',
                    description: 'Analyze competitor reaction scenarios',
                    variables: ['competitor_pricing', 'market_response', 'customer_loyalty'],
                    scenarios: ['No Response', 'Moderate Response', 'Aggressive Response']
                }
            ]
        };

        let result = templates;
        
        if (category !== 'all' && templates[category]) {
            result = { [category]: templates[category] };
        }

        // Industry-specific filtering
        if (industry) {
            const industryTemplates = WhatIfAnalysis.getIndustryTemplates(industry);
            result = { ...result, industry: industryTemplates };
        }

        res.json({
            success: true,
            templates: result,
            categories: Object.keys(templates),
            customization: {
                steps: [
                    'Select appropriate template',
                    'Customize variables for your business',
                    'Define scenario parameters',
                    'Set up calculations and dependencies',
                    'Create summary dashboard'
                ],
                bestPractices: [
                    'Start with simple scenarios',
                    'Document all assumptions',
                    'Use sensitivity analysis',
                    'Regular scenario updates',
                    'Validate with historical data'
                ]
            },
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Scenario Optimization
router.post('/optimize', async (req, res, next) => {
    try {
        const { scenarios, objective, constraints } = req.body;
        
        if (!scenarios || !objective) {
            return res.status(400).json({
                success: false,
                error: 'Scenarios and objective are required'
            });
        }

        // Optimize scenario parameters
        const optimization = WhatIfAnalysis.optimizeScenario(scenarios, objective, constraints);
        
        // Generate optimal scenario
        const optimalScenario = {
            name: 'Optimized Scenario',
            description: 'AI-optimized parameters for best outcome',
            variables: optimization.optimalVariables,
            expectedOutcome: optimization.optimalValue,
            confidence: optimization.confidence,
            sensitivity: optimization.sensitivity
        };

        res.json({
            success: true,
            optimalScenario,
            optimization,
            implementation: {
                solver: {
                    tool: 'Excel Solver',
                    steps: [
                        'Go to Data > Solver',
                        'Set objective cell',
                        'Define variable cells',
                        'Add constraints',
                        'Run optimization'
                    ]
                },
                validation: [
                    'Test optimal parameters',
                    'Verify constraint satisfaction',
                    'Perform sensitivity analysis',
                    'Document optimization results'
                ]
            },
            alternatives: optimization.alternatives || [],
            recommendations: [
                'Implement optimal scenario gradually',
                'Monitor key performance indicators',
                'Prepare for scenario adjustments',
                'Consider robustness over pure optimization'
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

module.exports = router;
