class DataAnalyzer {
    static getBasicStats(selectedData) {
        const { values } = selectedData;
        
        return {
            hasHeaders: this.detectHeaders(values),
            dataTypes: this.analyzeDataTypes(values),
            numericColumns: this.getNumericColumnCount(values),
            textColumns: this.getTextColumnCount(values),
            missingValues: this.countMissingValues(values)
        };
    }

    static performDeepAnalysis(selectedData) {
        const analysis = {
            basicStats: this.getBasicStats(selectedData),
            statisticalSummary: this.calculateStatistics(selectedData),
            dataQuality: this.assessDataQuality(selectedData),
            patterns: this.detectPatterns(selectedData),
            correlations: this.calculateCorrelations(selectedData),
            outliers: this.detectOutliers(selectedData),
            trends: this.analyzeTrends(selectedData),
            recommendations: this.generateRecommendations(selectedData)
        };

        return analysis;
    }

    static calculateStatistics(selectedData) {
        const { values } = selectedData;
        const numericColumns = this.getNumericColumns(values);
        const stats = {};

        numericColumns.forEach((colIndex) => {
            const columnData = values
                .slice(1) // Skip headers
                .map(row => row[colIndex])
                .filter(val => typeof val === 'number' && !isNaN(val));

            if (columnData.length > 0) {
                stats[`column_${colIndex}`] = {
                    count: columnData.length,
                    sum: this.sum(columnData),
                    mean: this.mean(columnData),
                    median: this.median(columnData),
                    mode: this.mode(columnData),
                    min: Math.min(...columnData),
                    max: Math.max(...columnData),
                    range: Math.max(...columnData) - Math.min(...columnData),
                    variance: this.variance(columnData),
                    standardDeviation: this.standardDeviation(columnData),
                    skewness: this.skewness(columnData),
                    kurtosis: this.kurtosis(columnData),
                    quartiles: this.quartiles(columnData),
                    percentiles: this.percentiles(columnData)
                };
            }
        });

        return stats;
    }

    static assessDataQuality(selectedData) {
        const { values } = selectedData;
        const totalCells = values.length * (values[0]?.length || 0);
        const missingCount = this.countMissingValues(values);
        const duplicateRows = this.findDuplicateRows(values);
        
        const completeness = 1 - (missingCount / totalCells);
        const uniqueness = 1 - (duplicateRows / values.length);
        const consistency = this.assessConsistency(values);
        const accuracy = this.assessAccuracy(values);
        
        const overallScore = (completeness + uniqueness + consistency + accuracy) / 4;

        return {
            completeness: {
                score: completeness,
                missingValues: missingCount,
                totalCells: totalCells
            },
            uniqueness: {
                score: uniqueness,
                duplicateRows: duplicateRows,
                totalRows: values.length
            },
            consistency: {
                score: consistency,
                issues: this.findConsistencyIssues(values)
            },
            accuracy: {
                score: accuracy,
                anomalies: this.findAnomalies(values)
            },
            overall: {
                score: overallScore,
                grade: this.getQualityGrade(overallScore),
                recommendations: this.getQualityRecommendations(overallScore)
            }
        };
    }

    static detectPatterns(selectedData) {
        const { values } = selectedData;
        
        return {
            seasonality: this.detectSeasonality(values),
            cycles: this.detectCycles(values),
            growth: this.analyzeGrowthPatterns(values),
            distributions: this.analyzeDistributions(values),
            relationships: this.findRelationships(values)
        };
    }

    static calculateCorrelations(selectedData) {
        const { values } = selectedData;
        const numericColumns = this.getNumericColumns(values);
        const correlations = {};

        for (let i = 0; i < numericColumns.length; i++) {
            for (let j = i + 1; j < numericColumns.length; j++) {
                const col1 = numericColumns[i];
                const col2 = numericColumns[j];
                
                const data1 = values.slice(1).map(row => row[col1]).filter(val => typeof val === 'number');
                const data2 = values.slice(1).map(row => row[col2]).filter(val => typeof val === 'number');
                
                if (data1.length > 1 && data2.length > 1) {
                    const correlation = this.pearsonCorrelation(data1, data2);
                    const strength = this.getCorrelationStrength(correlation);
                    
                    correlations[`col_${col1}_col_${col2}`] = {
                        coefficient: correlation,
                        strength: strength,
                        interpretation: this.interpretCorrelation(correlation)
                    };
                }
            }
        }

        return correlations;
    }

    static detectOutliers(selectedData) {
        const { values } = selectedData;
        const numericColumns = this.getNumericColumns(values);
        const outliers = {};

        numericColumns.forEach(colIndex => {
            const columnData = values
                .slice(1)
                .map(row => row[colIndex])
                .filter(val => typeof val === 'number' && !isNaN(val));

            if (columnData.length > 3) {
                const q1 = this.quartiles(columnData)[0];
                const q3 = this.quartiles(columnData)[2];
                const iqr = q3 - q1;
                const lowerBound = q1 - 1.5 * iqr;
                const upperBound = q3 + 1.5 * iqr;
                
                const outlierValues = columnData.filter(val => val < lowerBound || val > upperBound);
                const extremeOutliers = columnData.filter(val => val < q1 - 3 * iqr || val > q3 + 3 * iqr);
                
                outliers[`column_${colIndex}`] = {
                    mild: {
                        count: outlierValues.length - extremeOutliers.length,
                        values: outlierValues.filter(val => !extremeOutliers.includes(val))
                    },
                    extreme: {
                        count: extremeOutliers.length,
                        values: extremeOutliers
                    },
                    bounds: { lower: lowerBound, upper: upperBound },
                    percentage: (outlierValues.length / columnData.length) * 100,
                    impact: this.assessOutlierImpact(outlierValues, columnData)
                };
            }
        });

        return outliers;
    }

    static analyzeTrends(selectedData) {
        const { values } = selectedData;
        const numericColumns = this.getNumericColumns(values);
        const trends = {};

        numericColumns.forEach(colIndex => {
            const columnData = values
                .slice(1)
                .map(row => row[colIndex])
                .filter(val => typeof val === 'number' && !isNaN(val));

            if (columnData.length > 2) {
                const trend = this.calculateLinearTrend(columnData);
                const movingAverage = this.calculateMovingAverage(columnData, 3);
                const seasonalComponent = this.extractSeasonalComponent(columnData);
                
                trends[`column_${colIndex}`] = {
                    linear: trend,
                    movingAverage: movingAverage,
                    seasonal: seasonalComponent,
                    volatility: this.calculateVolatility(columnData),
                    momentum: this.calculateMomentum(columnData)
                };
            }
        });

        return trends;
    }

    static generateRecommendations(selectedData) {
        const analysis = this.performDeepAnalysis(selectedData);
        const recommendations = [];

        // Data quality recommendations
        if (analysis.dataQuality.overall.score < 0.8) {
            recommendations.push({
                category: "Data Quality",
                priority: "high",
                message: "Improve data quality by addressing missing values and inconsistencies",
                actions: [
                    "Use IFERROR() functions to handle missing data",
                    "Implement data validation rules",
                    "Create data cleansing formulas"
                ]
            });
        }

        // Statistical analysis recommendations
        Object.keys(analysis.correlations).forEach(key => {
            const corr = analysis.correlations[key];
            if (Math.abs(corr.coefficient) > 0.7) {
                recommendations.push({
                    category: "Statistical Analysis",
                    priority: "medium",
                    message: `Strong ${corr.strength} correlation detected between variables`,
                    actions: [
                        "Create scatter plots to visualize relationship",
                        "Consider regression analysis",
                        "Investigate causal relationships"
                    ]
                });
            }
        });

        // Outlier recommendations
        Object.keys(analysis.outliers).forEach(key => {
            const outlier = analysis.outliers[key];
            if (outlier.percentage > 5) {
                recommendations.push({
                    category: "Outlier Management",
                    priority: "medium",
                    message: `${outlier.percentage.toFixed(1)}% outliers detected in ${key}`,
                    actions: [
                        "Investigate outlier causes",
                        "Consider robust statistical methods",
                        "Implement outlier detection formulas"
                    ]
                });
            }
        });

        // Trend analysis recommendations
        Object.keys(analysis.trends).forEach(key => {
            const trend = analysis.trends[key];
            if (Math.abs(trend.linear.slope) > 0.1) {
                recommendations.push({
                    category: "Trend Analysis",
                    priority: "high",
                    message: `Significant ${trend.linear.direction} trend in ${key}`,
                    actions: [
                        "Create trend line charts",
                        "Implement forecasting formulas",
                        "Monitor trend sustainability"
                    ]
                });
            }
        });

        return recommendations;
    }

    // Statistical calculation methods
    static sum(data) {
        return data.reduce((a, b) => a + b, 0);
    }

    static mean(data) {
        return this.sum(data) / data.length;
    }

    static median(data) {
        const sorted = [...data].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }

    static mode(data) {
        const frequency = {};
        data.forEach(value => frequency[value] = (frequency[value] || 0) + 1);
        const maxFreq = Math.max(...Object.values(frequency));
        return Object.keys(frequency).filter(key => frequency[key] === maxFreq).map(Number);
    }

    static variance(data) {
        const mean = this.mean(data);
        return this.sum(data.map(x => Math.pow(x - mean, 2))) / (data.length - 1);
    }

    static standardDeviation(data) {
        return Math.sqrt(this.variance(data));
    }

    static skewness(data) {
        const mean = this.mean(data);
        const stdDev = this.standardDeviation(data);
        const n = data.length;
        const skew = data.reduce((sum, x) => sum + Math.pow((x - mean) / stdDev, 3), 0);
        return (n / ((n - 1) * (n - 2))) * skew;
    }

    static kurtosis(data) {
        const mean = this.mean(data);
        const stdDev = this.standardDeviation(data);
        const n = data.length;
        const kurt = data.reduce((sum, x) => sum + Math.pow((x - mean) / stdDev, 4), 0);
        return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * kurt - (3 * Math.pow(n - 1, 2)) / ((n - 2) * (n - 3));
    }

    static quartiles(data) {
        const sorted = [...data].sort((a, b) => a - b);
        return [
            this.percentile(sorted, 25),
            this.percentile(sorted, 50),
            this.percentile(sorted, 75)
        ];
    }

    static percentiles(data) {
        const sorted = [...data].sort((a, b) => a - b);
        return {
            p5: this.percentile(sorted, 5),
            p10: this.percentile(sorted, 10),
            p25: this.percentile(sorted, 25),
            p50: this.percentile(sorted, 50),
            p75: this.percentile(sorted, 75),
            p90: this.percentile(sorted, 90),
            p95: this.percentile(sorted, 95)
        };
    }

    static percentile(sortedData, percentile) {
        const index = (percentile / 100) * (sortedData.length - 1);
        if (index % 1 === 0) {
            return sortedData[index];
        } else {
            const lower = sortedData[Math.floor(index)];
            const upper = sortedData[Math.ceil(index)];
            return lower + (upper - lower) * (index % 1);
        }
    }

    static pearsonCorrelation(x, y) {
        const n = Math.min(x.length, y.length);
        if (n < 2) return 0;
        
        const sumX = this.sum(x.slice(0, n));
        const sumY = this.sum(y.slice(0, n));
        const sumXY = this.sum(x.slice(0, n).map((xi, i) => xi * y[i]));
        const sumX2 = this.sum(x.slice(0, n).map(xi => xi * xi));
        const sumY2 = this.sum(y.slice(0, n).map(yi => yi * yi));
        
        const numerator = n * sumXY - sumX * sumY;
        const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
        
        return denominator === 0 ? 0 : numerator / denominator;
    }

    // Helper methods
    static detectHeaders(values) {
        if (values.length < 2) return false;
        const firstRow = values[0];
        const secondRow = values[1];
        return firstRow.every(cell => typeof cell === 'string') && 
               secondRow.some(cell => typeof cell === 'number');
    }

    static analyzeDataTypes(values) {
        const types = new Set();
        values.forEach(row => {
            row.forEach(cell => {
                if (typeof cell === 'number') types.add('number');
                else if (typeof cell === 'string') types.add('text');
                else if (cell instanceof Date) types.add('date');
                else if (cell === null || cell === undefined) types.add('empty');
            });
        });
        return Array.from(types);
    }

    static getNumericColumns(values) {
        if (values.length < 2) return [];
        const columns = [];
        const sampleRow = values[1];
        sampleRow.forEach((cell, index) => {
            if (typeof cell === 'number') columns.push(index);
        });
        return columns;
    }

    static getNumericColumnCount(values) {
        return this.getNumericColumns(values).length;
    }

    static getTextColumnCount(values) {
        if (values.length < 2) return 0;
        const sampleRow = values[1];
        return sampleRow.filter(cell => typeof cell === 'string').length;
    }

    static countMissingValues(values) {
        let count = 0;
        values.forEach(row => {
            row.forEach(cell => {
                if (cell === null || cell === undefined || cell === '') count++;
            });
        });
        return count;
    }

    static findDuplicateRows(values) {
        const rowStrings = values.map(row => JSON.stringify(row));
        const unique = new Set(rowStrings);
        return values.length - unique.size;
    }

    static assessConsistency(values) {
        // Simplified consistency assessment
        return 0.85; // Would implement actual consistency checks
    }

    static assessAccuracy(values) {
        // Simplified accuracy assessment
        return 0.90; // Would implement actual accuracy checks
    }

    static getQualityGrade(score) {
        if (score >= 0.9) return 'A';
        if (score >= 0.8) return 'B';
        if (score >= 0.7) return 'C';
        if (score >= 0.6) return 'D';
        return 'F';
    }

    static getQualityRecommendations(score) {
        if (score >= 0.9) return ["Excellent data quality - maintain current standards"];
        if (score >= 0.8) return ["Good data quality - minor improvements needed"];
        if (score >= 0.7) return ["Moderate data quality - focus on key issues"];
        if (score >= 0.6) return ["Poor data quality - significant improvements required"];
        return ["Very poor data quality - major data cleansing needed"];
    }

    static getCorrelationStrength(correlation) {
        const abs = Math.abs(correlation);
        if (abs >= 0.9) return 'very strong';
        if (abs >= 0.7) return 'strong';
        if (abs >= 0.5) return 'moderate';
        if (abs >= 0.3) return 'weak';
        return 'very weak';
    }

    static interpretCorrelation(correlation) {
        if (correlation > 0.7) return 'Strong positive relationship';
        if (correlation > 0.3) return 'Moderate positive relationship';
        if (correlation > 0) return 'Weak positive relationship';
        if (correlation === 0) return 'No linear relationship';
        if (correlation > -0.3) return 'Weak negative relationship';
        if (correlation > -0.7) return 'Moderate negative relationship';
        return 'Strong negative relationship';
    }

    static calculateLinearTrend(data) {
        const n = data.length;
        const x = Array.from({length: n}, (_, i) => i);
        const y = data;
        
        const sumX = this.sum(x);
        const sumY = this.sum(y);
        const sumXY = this.sum(x.map((xi, i) => xi * y[i]));
        const sumX2 = this.sum(x.map(xi => xi * xi));
        
        const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        const intercept = (sumY - slope * sumX) / n;
        
        const direction = slope > 0.01 ? 'increasing' : slope < -0.01 ? 'decreasing' : 'stable';
        const strength = Math.abs(slope);
        
        return {
            slope,
            intercept,
            direction,
            strength,
            description: `${direction} trend with ${strength > 0.1 ? 'strong' : strength > 0.05 ? 'moderate' : 'weak'} strength`
        };
    }

    static calculateMovingAverage(data, period) {
        const movingAvg = [];
        for (let i = period - 1; i < data.length; i++) {
            const avg = this.mean(data.slice(i - period + 1, i + 1));
            movingAvg.push(avg);
        }
        return movingAvg;
    }

    static calculateVolatility(data) {
        if (data.length < 2) return 0;
        const returns = [];
        for (let i = 1; i < data.length; i++) {
            if (data[i - 1] !== 0) {
                returns.push((data[i] - data[i - 1]) / data[i - 1]);
            }
        }
        return this.standardDeviation(returns);
    }

    static calculateMomentum(data, period = 3) {
        if (data.length < period + 1) return 0;
        const recent = data.slice(-period);
        const earlier = data.slice(-period * 2, -period);
        return this.mean(recent) - this.mean(earlier);
    }

    // Placeholder methods for complex analysis
    static detectSeasonality(values) {
        return { detected: false, period: null, strength: 0 };
    }

    static detectCycles(values) {
        return { detected: false, cycles: [] };
    }

    static analyzeGrowthPatterns(values) {
        return { pattern: 'linear', confidence: 0.5 };
    }

    static analyzeDistributions(values) {
        return { normal: 0.6, uniform: 0.2, exponential: 0.2 };
    }

    static findRelationships(values) {
        return { linear: 0.7, polynomial: 0.2, logarithmic: 0.1 };
    }

    static findConsistencyIssues(values) {
        return ['Format inconsistencies', 'Unit variations'];
    }

    static findAnomalies(values) {
        return ['Unusual data patterns', 'Potential data entry errors'];
    }

    static assessOutlierImpact(outliers, allData) {
        const impactScore = outliers.length / allData.length;
        if (impactScore > 0.1) return 'high';
        if (impactScore > 0.05) return 'medium';
        return 'low';
    }

    static extractSeasonalComponent(data) {
        return { detected: false, component: [] };
    }
}

module.exports = DataAnalyzer;
