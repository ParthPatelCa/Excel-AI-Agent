class FormulaGenerator {
    static generateFormula(intent, dataContext, parameters = {}) {
        const formulaMap = {
            // Basic arithmetic
            sum: this.generateSumFormula,
            average: this.generateAverageFormula,
            count: this.generateCountFormula,
            max: this.generateMaxFormula,
            min: this.generateMinFormula,
            
            // Conditional formulas
            sumif: this.generateSumIfFormula,
            countif: this.generateCountIfFormula,
            averageif: this.generateAverageIfFormula,
            
            // Advanced statistical
            correlation: this.generateCorrelationFormula,
            regression: this.generateRegressionFormula,
            forecast: this.generateForecastFormula,
            trend: this.generateTrendFormula,
            
            // Financial formulas
            npv: this.generateNPVFormula,
            irr: this.generateIRRFormula,
            pmt: this.generatePMTFormula,
            fv: this.generateFVFormula,
            
            // What-if analysis
            goalseek: this.generateGoalSeekFormula,
            sensitivity: this.generateSensitivityFormula,
            scenario: this.generateScenarioFormula,
            
            // Data analysis
            pivot: this.generatePivotFormula,
            subtotal: this.generateSubtotalFormula,
            rank: this.generateRankFormula,
            percentile: this.generatePercentileFormula
        };

        const generator = formulaMap[intent];
        if (generator) {
            return generator.call(this, dataContext, parameters);
        }

        return this.generateCustomFormula(intent, dataContext, parameters);
    }

    // Basic Formula Generators
    static generateSumFormula(dataContext, parameters) {
        const range = parameters.range || dataContext.address;
        return {
            formula: `=SUM(${range})`,
            description: `Sum all values in range ${range}`,
            category: 'Basic Math',
            alternatives: [
                `=SUMPRODUCT(${range})`,
                `=AGGREGATE(9,5,${range})` // Ignore errors
            ]
        };
    }

    static generateAverageFormula(dataContext, parameters) {
        const range = parameters.range || dataContext.address;
        return {
            formula: `=AVERAGE(${range})`,
            description: `Calculate the mean of values in ${range}`,
            category: 'Statistics',
            alternatives: [
                `=AGGREGATE(1,5,${range})`, // Ignore errors
                `=TRIMMEAN(${range},0.1)`, // Exclude top/bottom 5%
                `=MEDIAN(${range})` // Alternative central tendency
            ]
        };
    }

    static generateCountFormula(dataContext, parameters) {
        const range = parameters.range || dataContext.address;
        const countType = parameters.type || 'numbers';
        
        const formulas = {
            numbers: `=COUNT(${range})`,
            nonEmpty: `=COUNTA(${range})`,
            empty: `=COUNTBLANK(${range})`,
            unique: `=SUMPRODUCT(1/COUNTIF(${range},${range}))`
        };

        return {
            formula: formulas[countType],
            description: `Count ${countType} in range ${range}`,
            category: 'Counting',
            alternatives: Object.entries(formulas)
                .filter(([key]) => key !== countType)
                .map(([key, formula]) => `${formula} (${key})`)
        };
    }

    // Conditional Formula Generators
    static generateSumIfFormula(dataContext, parameters) {
        const range = parameters.range || dataContext.address;
        const criteria = parameters.criteria || '>0';
        
        return {
            formula: `=SUMIF(${range},"${criteria}")`,
            description: `Sum values in ${range} that meet criteria: ${criteria}`,
            category: 'Conditional',
            variations: [
                `=SUMIFS(${range},${range},"${criteria}")`, // Multiple criteria version
                `=SUMPRODUCT((${range}${criteria.replace(/[<>=]/g, '')})*(${range}))` // Array formula
            ],
            examples: [
                'SUMIF(A1:A10,">100") - Sum values greater than 100',
                'SUMIF(A1:A10,"<>0") - Sum non-zero values',
                'SUMIF(B1:B10,"Product1",A1:A10) - Sum A values where B="Product1"'
            ]
        };
    }

    static generateCountIfFormula(dataContext, parameters) {
        const range = parameters.range || dataContext.address;
        const criteria = parameters.criteria || '>0';
        
        return {
            formula: `=COUNTIF(${range},"${criteria}")`,
            description: `Count cells in ${range} that meet criteria: ${criteria}`,
            category: 'Conditional Counting',
            variations: [
                `=COUNTIFS(${range},"${criteria}")`, // Multiple criteria
                `=SUMPRODUCT(--(${range}${criteria.replace(/[<>=]/g, '')}))` // Array formula
            ]
        };
    }

    // Advanced Statistical Formulas
    static generateCorrelationFormula(dataContext, parameters) {
        const range1 = parameters.range1 || 'A:A';
        const range2 = parameters.range2 || 'B:B';
        
        return {
            formula: `=CORREL(${range1},${range2})`,
            description: `Calculate correlation coefficient between ${range1} and ${range2}`,
            category: 'Statistical Analysis',
            interpretation: {
                '>0.7': 'Strong positive correlation',
                '0.3-0.7': 'Moderate positive correlation',
                '-0.3-0.3': 'Weak or no correlation',
                '-0.7--0.3': 'Moderate negative correlation',
                '<-0.7': 'Strong negative correlation'
            },
            relatedFormulas: [
                `=RSQ(${range1},${range2})` + ' (R-squared)',
                `=COVARIANCE.P(${range1},${range2})` + ' (Covariance)'
            ]
        };
    }

    static generateRegressionFormula(dataContext, parameters) {
        const yRange = parameters.yRange || 'A:A';
        const xRange = parameters.xRange || 'B:B';
        
        return {
            formula: `=LINEST(${yRange},${xRange},TRUE,TRUE)`,
            description: `Linear regression analysis for ${yRange} vs ${xRange}`,
            category: 'Predictive Analysis',
            components: {
                slope: `=INDEX(LINEST(${yRange},${xRange}),1,1)`,
                intercept: `=INDEX(LINEST(${yRange},${xRange}),1,2)`,
                rSquared: `=INDEX(LINEST(${yRange},${xRange},TRUE,TRUE),3,1)`
            },
            prediction: `=INDEX(LINEST(${yRange},${xRange}),1,2) + INDEX(LINEST(${yRange},${xRange}),1,1)*NEW_X_VALUE`
        };
    }

    static generateForecastFormula(dataContext, parameters) {
        const xValue = parameters.xValue || 'NEW_X';
        const yRange = parameters.yRange || 'A:A';
        const xRange = parameters.xRange || 'B:B';
        
        return {
            formula: `=FORECAST(${xValue},${yRange},${xRange})`,
            description: `Forecast Y value for X=${xValue} based on historical data`,
            category: 'Forecasting',
            alternatives: [
                `=FORECAST.LINEAR(${xValue},${yRange},${xRange})`, // Excel 2016+
                `=TREND(${yRange},${xRange},${xValue})`, // Trend-based forecast
                `=GROWTH(${yRange},${xRange},${xValue})` // Exponential growth
            ],
            confidence: `=FORECAST.LINEAR.CONFI(0.05,STEYX(${yRange},${xRange}),COUNT(${xRange}),AVERAGE(${xRange}),DEVSQ(${xRange}))`
        };
    }

    // Financial Formulas
    static generateNPVFormula(dataContext, parameters) {
        const rate = parameters.rate || '10%';
        const cashFlows = parameters.cashFlows || 'B2:B10';
        
        return {
            formula: `=NPV(${rate},${cashFlows})`,
            description: `Net Present Value at ${rate} discount rate`,
            category: 'Financial Analysis',
            interpretation: {
                '>0': 'Positive NPV - Investment is profitable',
                '=0': 'Break-even point',
                '<0': 'Negative NPV - Investment loses money'
            },
            relatedMetrics: {
                irr: `=IRR(${cashFlows})`,
                payback: `Custom calculation needed`,
                profitabilityIndex: `=NPV(${rate},${cashFlows})/ABS(B1)`
            }
        };
    }

    static generateIRRFormula(dataContext, parameters) {
        const cashFlows = parameters.cashFlows || 'B1:B10';
        const guess = parameters.guess || '10%';
        
        return {
            formula: `=IRR(${cashFlows},${guess})`,
            description: `Internal Rate of Return for cash flow series`,
            category: 'Financial Analysis',
            validation: `=IF(IRR(${cashFlows})>WACC,"Accept","Reject")`,
            sensitivity: `=IRR(${cashFlows}*(1+SENSITIVITY_FACTOR))`
        };
    }

    // What-If Analysis Formulas
    static generateGoalSeekFormula(dataContext, parameters) {
        const targetCell = parameters.targetCell || 'C1';
        const targetValue = parameters.targetValue || '100';
        const changingCell = parameters.changingCell || 'A1';
        
        return {
            formula: `Goal Seek Setup Required`,
            description: `Find ${changingCell} value to make ${targetCell} equal ${targetValue}`,
            category: 'What-If Analysis',
            steps: [
                `1. Go to Data > What-If Analysis > Goal Seek`,
                `2. Set cell: ${targetCell}`,
                `3. To value: ${targetValue}`,
                `4. By changing cell: ${changingCell}`
            ],
            formula_alternative: `=SOLVER_SETUP("${targetCell}","${targetValue}","${changingCell}")`
        };
    }

    static generateSensitivityFormula(dataContext, parameters) {
        const inputCell = parameters.inputCell || 'A1';
        const outputCell = parameters.outputCell || 'B1';
        const changePercent = parameters.changePercent || '10%';
        
        return {
            formula: `=${outputCell}*(1+${changePercent})`,
            description: `Sensitivity analysis: ${changePercent} change in ${inputCell}`,
            category: 'Sensitivity Analysis',
            dataTable: {
                setup: `Create data table with ${inputCell} variations`,
                formula: `=${outputCell}`,
                instructions: [
                    '1. Create column with input variations',
                    '2. Reference output formula in adjacent cell',
                    '3. Select range and use Data > What-If Analysis > Data Table'
                ]
            }
        };
    }

    // Advanced Analysis Formulas
    static generatePivotFormula(dataContext, parameters) {
        const sourceRange = parameters.sourceRange || dataContext.address;
        
        return {
            formula: `=GETPIVOTDATA("Sum of Value",PivotTable,"Field1","Criteria")`,
            description: `Extract data from PivotTable dynamically`,
            category: 'Data Analysis',
            setup: [
                '1. Create PivotTable from source data',
                '2. Configure rows, columns, and values',
                '3. Use GETPIVOTDATA to reference results'
            ],
            alternatives: [
                'Manual PivotTable creation',
                'Power Query for data transformation',
                'SUMIFS/COUNTIFS for simple aggregations'
            ]
        };
    }

    static generateSubtotalFormula(dataContext, parameters) {
        const functionNum = parameters.functionNum || '9'; // SUM
        const range = parameters.range || dataContext.address;
        
        const functions = {
            '1': 'AVERAGE',
            '2': 'COUNT',
            '3': 'COUNTA',
            '4': 'MAX',
            '5': 'MIN',
            '6': 'PRODUCT',
            '7': 'STDEV',
            '8': 'STDEVP',
            '9': 'SUM',
            '10': 'VAR',
            '11': 'VARP'
        };
        
        return {
            formula: `=SUBTOTAL(${functionNum},${range})`,
            description: `${functions[functionNum]} ignoring hidden rows in ${range}`,
            category: 'Data Analysis',
            functions: functions,
            useCase: 'Works with filtered data and ignores hidden rows'
        };
    }

    // Utility Methods
    static extractFormula(aiResponse) {
        // Extract formulas from AI response
        const formulaPattern = /=([A-Z0-9+\-*/(),.:\s$"']+)/gi;
        const matches = aiResponse.match(formulaPattern);
        
        if (matches && matches.length > 0) {
            const formula = matches[0];
            
            if (this.validateFormula(formula)) {
                return {
                    type: 'formula',
                    formula: formula,
                    buttonText: 'Insert Formula',
                    description: this.extractFormulaDescription(aiResponse, formula)
                };
            }
        }
        
        return null;
    }

    static validateFormula(formula) {
        // Basic Excel formula validation
        const validFunctions = [
            'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP',
            'INDEX', 'MATCH', 'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN',
            'SUMIF', 'COUNTIF', 'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS',
            'NPV', 'IRR', 'PMT', 'FV', 'PV', 'RATE', 'NPER',
            'CORREL', 'LINEST', 'FORECAST', 'TREND', 'GROWTH',
            'SUBTOTAL', 'GETPIVOTDATA', 'TRANSPOSE', 'SORT', 'FILTER'
        ];
        
        // Check for dangerous patterns
        const dangerousPatterns = [
            /\b(DELETE|DROP|EXEC|SHELL|KILL|SYSTEM)\b/i,
            /javascript:/i,
            /<script/i
        ];
        
        for (const pattern of dangerousPatterns) {
            if (pattern.test(formula)) {
                return false;
            }
        }
        
        // Extract function names
        const functionMatches = formula.match(/[A-Z]+(?=\s*\()/g);
        if (functionMatches) {
            const invalidFunctions = functionMatches.filter(func => 
                !validFunctions.includes(func.toUpperCase())
            );
            return invalidFunctions.length === 0;
        }
        
        return true;
    }

    static extractFormulaDescription(response, formula) {
        const lines = response.split('\n');
        const formulaIndex = lines.findIndex(line => line.includes(formula));
        
        if (formulaIndex !== -1) {
            // Look for description in surrounding lines
            for (let i = Math.max(0, formulaIndex - 2); i <= Math.min(lines.length - 1, formulaIndex + 2); i++) {
                const line = lines[i].trim();
                if (line && !line.includes(formula) && !line.startsWith('=')) {
                    return line;
                }
            }
        }
        
        return 'AI-generated formula';
    }

    static generateCustomFormula(intent, dataContext, parameters) {
        // Fallback for custom or complex formulas
        return {
            formula: '=CUSTOM_FORMULA_NEEDED',
            description: `Custom formula for: ${intent}`,
            category: 'Custom',
            recommendation: 'Provide more specific requirements for formula generation'
        };
    }

    static getFormulaExamples() {
        return {
            basic: [
                '=SUM(A1:A10) - Sum range of values',
                '=AVERAGE(B1:B10) - Calculate mean',
                '=COUNT(C1:C10) - Count numeric values'
            ],
            conditional: [
                '=SUMIF(A1:A10,">100") - Sum values greater than 100',
                '=COUNTIF(B1:B10,"Product") - Count specific text',
                '=AVERAGEIF(C1:C10,">0") - Average positive values'
            ],
            advanced: [
                '=VLOOKUP(E1,A1:C10,2,FALSE) - Lookup value',
                '=INDEX(MATCH(E1,A1:A10,0),B1:B10) - Advanced lookup',
                '=CORREL(A1:A10,B1:B10) - Correlation coefficient'
            ],
            financial: [
                '=NPV(10%,B2:B10) - Net present value',
                '=IRR(A1:A10) - Internal rate of return',
                '=PMT(5%/12,60,10000) - Loan payment'
            ]
        };
    }
}

module.exports = FormulaGenerator;
