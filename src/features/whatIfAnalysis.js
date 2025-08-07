class WhatIfAnalysis {
    static generateScenarios(selectedData) {
        const scenarios = {
            optimistic: this.createOptimisticScenario(selectedData),
            pessimistic: this.createPessimisticScenario(selectedData),
            realistic: this.createRealisticScenario(selectedData),
            custom: this.createCustomScenarios(selectedData)
        };

        return scenarios;
    }

    static createOptimisticScenario(selectedData) {
        const numericColumns = this.findNumericColumns(selectedData);
        
        return {
            name: "Optimistic Scenario",
            description: "Best-case projections with positive growth assumptions",
            assumptions: [
                "15-25% growth in key metrics",
                "Improved operational efficiency",
                "Market expansion opportunities"
            ],
            modifications: numericColumns.map(col => ({
                column: col,
                change: "+20%",
                formula: `=ORIGINAL_VALUE * 1.2`,
                rationale: "Optimistic growth projection"
            })),
            expectedOutcome: "Significant improvement in overall performance",
            riskLevel: "Low",
            probability: "25%"
        };
    }

    static createPessimisticScenario(selectedData) {
        const numericColumns = this.findNumericColumns(selectedData);
        
        return {
            name: "Pessimistic Scenario",
            description: "Worst-case projections with conservative assumptions",
            assumptions: [
                "10-15% decline in key metrics",
                "Economic downturn impact",
                "Increased operational costs"
            ],
            modifications: numericColumns.map(col => ({
                column: col,
                change: "-15%",
                formula: `=ORIGINAL_VALUE * 0.85`,
                rationale: "Conservative downturn projection"
            })),
            expectedOutcome: "Challenging performance requiring mitigation strategies",
            riskLevel: "High",
            probability: "15%"
        };
    }

    static createRealisticScenario(selectedData) {
        const numericColumns = this.findNumericColumns(selectedData);
        
        return {
            name: "Realistic Scenario",
            description: "Most likely projections based on current trends",
            assumptions: [
                "5-8% steady growth",
                "Stable market conditions",
                "Incremental improvements"
            ],
            modifications: numericColumns.map(col => ({
                column: col,
                change: "+6%",
                formula: `=ORIGINAL_VALUE * 1.06`,
                rationale: "Realistic growth based on historical trends"
            })),
            expectedOutcome: "Steady, sustainable growth",
            riskLevel: "Medium",
            probability: "60%"
        };
    }

    static createCustomScenarios(selectedData) {
        return [
            {
                name: "Cost Reduction Focus",
                description: "Scenario focusing on operational efficiency",
                modifications: [
                    { type: "cost_reduction", percentage: 12, target: "operational_costs" },
                    { type: "efficiency_gain", percentage: 8, target: "productivity" }
                ]
            },
            {
                name: "Market Expansion",
                description: "Aggressive growth through market expansion",
                modifications: [
                    { type: "revenue_increase", percentage: 30, target: "sales" },
                    { type: "investment_increase", percentage: 25, target: "marketing" }
                ]
            },
            {
                name: "Technology Upgrade",
                description: "Impact of digital transformation",
                modifications: [
                    { type: "efficiency_gain", percentage: 20, target: "automation" },
                    { type: "cost_increase", percentage: 15, target: "technology_investment" }
                ]
            }
        ];
    }

    static performSensitivityAnalysis(selectedData, variable, changeRange = [-50, 50]) {
        const analysis = {
            variable: variable,
            changeRange: changeRange,
            impacts: [],
            recommendations: []
        };

        // Generate sensitivity data points
        for (let change = changeRange[0]; change <= changeRange[1]; change += 10) {
            const impact = this.calculateImpact(selectedData, variable, change);
            analysis.impacts.push({
                change: `${change}%`,
                value: impact.value,
                impact: impact.description,
                formula: impact.formula
            });
        }

        // Generate recommendations based on sensitivity
        analysis.recommendations = this.generateSensitivityRecommendations(analysis.impacts);

        return analysis;
    }

    static calculateImpact(selectedData, variable, changePercentage) {
        const multiplier = 1 + (changePercentage / 100);
        
        return {
            value: `${changePercentage > 0 ? '+' : ''}${changePercentage}%`,
            description: this.getImpactDescription(changePercentage),
            formula: `=ORIGINAL_${variable.toUpperCase()} * ${multiplier}`
        };
    }

    static getImpactDescription(changePercentage) {
        if (changePercentage > 20) return "Significant positive impact";
        if (changePercentage > 10) return "Moderate positive impact";
        if (changePercentage > 0) return "Minor positive impact";
        if (changePercentage === 0) return "No change";
        if (changePercentage > -10) return "Minor negative impact";
        if (changePercentage > -20) return "Moderate negative impact";
        return "Significant negative impact";
    }

    static generateSensitivityRecommendations(impacts) {
        const recommendations = [];
        
        // Find most sensitive ranges
        const highImpactChanges = impacts.filter(impact => 
            Math.abs(parseFloat(impact.change)) > 15
        );

        if (highImpactChanges.length > 0) {
            recommendations.push({
                type: "risk_management",
                message: "High sensitivity detected - implement monitoring and controls",
                priority: "high"
            });
        }

        recommendations.push({
            type: "optimization",
            message: "Focus on variables with highest positive impact potential",
            priority: "medium"
        });

        recommendations.push({
            type: "contingency",
            message: "Prepare contingency plans for negative scenarios",
            priority: "medium"
        });

        return recommendations;
    }

    static generateGoalSeekAnalysis(selectedData, targetValue, changingCell) {
        return {
            objective: `Reach target value of ${targetValue}`,
            changingCell: changingCell,
            currentValue: this.getCurrentValue(selectedData, changingCell),
            requiredChange: this.calculateRequiredChange(selectedData, targetValue, changingCell),
            formula: `=GOAL.SEEK(target_cell, ${targetValue}, ${changingCell})`,
            feasibility: this.assessFeasibility(selectedData, targetValue, changingCell),
            steps: [
                "1. Set up the formula linking target to changing cell",
                "2. Use Data > What-If Analysis > Goal Seek",
                "3. Set target value and changing cell",
                "4. Analyze the required change for feasibility"
            ]
        };
    }

    static createDataTable(selectedData, inputCells, resultCells) {
        return {
            type: "one_way_data_table",
            setup: {
                inputCell: inputCells[0],
                resultFormula: resultCells[0],
                inputValues: this.generateInputRange(selectedData, inputCells[0])
            },
            instructions: [
                "1. Create a column with different input values",
                "2. Create a formula referencing the result cell",
                "3. Select the range including both",
                "4. Use Data > What-If Analysis > Data Table",
                "5. Specify the input cell reference"
            ],
            expectedOutput: "Table showing how result changes with different inputs"
        };
    }

    static generateMonteCarloInputs(selectedData, iterations = 1000) {
        const inputs = {
            iterations: iterations,
            randomVariables: this.identifyRandomVariables(selectedData),
            distributions: this.suggestDistributions(selectedData),
            formulas: [
                "=RAND() for uniform distribution",
                "=NORM.INV(RAND(), mean, std_dev) for normal distribution",
                "=GAMMA.INV(RAND(), alpha, beta) for gamma distribution"
            ],
            setup: [
                "1. Identify uncertain variables",
                "2. Define probability distributions",
                "3. Create random number generators",
                "4. Link to result calculations",
                "5. Run multiple iterations",
                "6. Analyze statistical results"
            ]
        };

        return inputs;
    }

    // Helper methods
    static findNumericColumns(selectedData) {
        const numericColumns = [];
        if (selectedData.values.length > 1) {
            const dataRow = selectedData.values[1]; // Skip potential headers
            dataRow.forEach((cell, index) => {
                if (typeof cell === 'number' && !isNaN(cell)) {
                    numericColumns.push(`Column ${String.fromCharCode(65 + index)}`);
                }
            });
        }
        return numericColumns;
    }

    static getCurrentValue(selectedData, changingCell) {
        // Simplified - in real implementation, would parse cell reference
        return "Current value from " + changingCell;
    }

    static calculateRequiredChange(selectedData, targetValue, changingCell) {
        // Simplified calculation
        return `Change needed to reach ${targetValue}`;
    }

    static assessFeasibility(selectedData, targetValue, changingCell) {
        // Basic feasibility assessment
        return {
            feasible: true,
            confidence: "medium",
            risks: ["Market conditions", "Resource constraints"],
            recommendations: ["Gradual implementation", "Monitor key metrics"]
        };
    }

    static generateInputRange(selectedData, inputCell) {
        // Generate a range of values for sensitivity analysis
        return Array.from({length: 11}, (_, i) => (i - 5) * 10); // -50% to +50%
    }

    static identifyRandomVariables(selectedData) {
        return [
            "Sales volume (normal distribution)",
            "Price fluctuation (uniform distribution)",
            "Cost variations (gamma distribution)"
        ];
    }

    static suggestDistributions(selectedData) {
        return {
            normal: "For variables with symmetric variation around a mean",
            uniform: "For variables with equal probability within a range",
            gamma: "For skewed positive variables like costs or times",
            triangular: "For variables with minimum, most likely, and maximum values"
        };
    }
}

module.exports = WhatIfAnalysis;
