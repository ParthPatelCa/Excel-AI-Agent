const express = require('express');
const router = express.Router();
const FormulaGenerator = require('../features/formulaGenerator');

// Generate Formula
router.post('/generate', async (req, res, next) => {
    try {
        const { description, dataContext, formulaType = 'auto' } = req.body;
        
        if (!description) {
            return res.status(400).json({
                success: false,
                error: 'Formula description is required'
            });
        }

        const formula = FormulaGenerator.generateFormula(description, dataContext);

        res.json({
            success: true,
            formula,
            explanation: {
                description: formula.description,
                usage: formula.usage,
                examples: formula.examples
            },
            validation: FormulaGenerator.validateFormula(formula.formula),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Financial Formulas
router.post('/financial', async (req, res, next) => {
    try {
        const { type, parameters } = req.body;
        
        if (!type) {
            return res.status(400).json({
                success: false,
                error: 'Financial formula type is required'
            });
        }

        let formula;
        
        switch (type) {
            case 'npv':
                formula = FormulaGenerator.generateNPVFormula(parameters);
                break;
            case 'irr':
                formula = FormulaGenerator.generateIRRFormula(parameters);
                break;
            case 'pmt':
                formula = FormulaGenerator.generatePMTFormula(parameters);
                break;
            case 'fv':
                formula = FormulaGenerator.generateFVFormula(parameters);
                break;
            case 'pv':
                formula = FormulaGenerator.generatePVFormula(parameters);
                break;
            case 'rate':
                formula = FormulaGenerator.generateRateFormula(parameters);
                break;
            case 'nper':
                formula = FormulaGenerator.generateNPERFormula(parameters);
                break;
            default:
                return res.status(400).json({
                    success: false,
                    error: `Unsupported financial formula type: ${type}`
                });
        }

        res.json({
            success: true,
            formula,
            category: 'Financial',
            validation: FormulaGenerator.validateFormula(formula.formula),
            relatedFormulas: FormulaGenerator.getRelatedFormulas(type, 'financial'),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Statistical Formulas
router.post('/statistical', async (req, res, next) => {
    try {
        const { type, parameters } = req.body;
        
        if (!type) {
            return res.status(400).json({
                success: false,
                error: 'Statistical formula type is required'
            });
        }

        let formula;
        
        switch (type) {
            case 'correlation':
                formula = FormulaGenerator.generateCorrelationFormula(parameters);
                break;
            case 'regression':
                formula = FormulaGenerator.generateRegressionFormula(parameters);
                break;
            case 'ttest':
                formula = FormulaGenerator.generateTTestFormula(parameters);
                break;
            case 'anova':
                formula = FormulaGenerator.generateANOVAFormula(parameters);
                break;
            case 'descriptive':
                formula = FormulaGenerator.generateDescriptiveStatsFormula(parameters);
                break;
            case 'confidence':
                formula = FormulaGenerator.generateConfidenceIntervalFormula(parameters);
                break;
            default:
                return res.status(400).json({
                    success: false,
                    error: `Unsupported statistical formula type: ${type}`
                });
        }

        res.json({
            success: true,
            formula,
            category: 'Statistical',
            validation: FormulaGenerator.validateFormula(formula.formula),
            relatedFormulas: FormulaGenerator.getRelatedFormulas(type, 'statistical'),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Lookup Formulas
router.post('/lookup', async (req, res, next) => {
    try {
        const { type, parameters } = req.body;
        
        if (!type) {
            return res.status(400).json({
                success: false,
                error: 'Lookup formula type is required'
            });
        }

        let formula;
        
        switch (type) {
            case 'vlookup':
                formula = FormulaGenerator.generateVLOOKUPFormula(parameters);
                break;
            case 'hlookup':
                formula = FormulaGenerator.generateHLOOKUPFormula(parameters);
                break;
            case 'xlookup':
                formula = FormulaGenerator.generateXLOOKUPFormula(parameters);
                break;
            case 'index-match':
                formula = FormulaGenerator.generateIndexMatchFormula(parameters);
                break;
            case 'filter':
                formula = FormulaGenerator.generateFilterFormula(parameters);
                break;
            default:
                return res.status(400).json({
                    success: false,
                    error: `Unsupported lookup formula type: ${type}`
                });
        }

        res.json({
            success: true,
            formula,
            category: 'Lookup',
            validation: FormulaGenerator.validateFormula(formula.formula),
            relatedFormulas: FormulaGenerator.getRelatedFormulas(type, 'lookup'),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Conditional Formulas
router.post('/conditional', async (req, res, next) => {
    try {
        const { type, parameters } = req.body;
        
        if (!type) {
            return res.status(400).json({
                success: false,
                error: 'Conditional formula type is required'
            });
        }

        let formula;
        
        switch (type) {
            case 'if':
                formula = FormulaGenerator.generateIfFormula(parameters);
                break;
            case 'nested-if':
                formula = FormulaGenerator.generateNestedIfFormula(parameters);
                break;
            case 'ifs':
                formula = FormulaGenerator.generateIfsFormula(parameters);
                break;
            case 'sumif':
                formula = FormulaGenerator.generateSumIfFormula(parameters);
                break;
            case 'countif':
                formula = FormulaGenerator.generateCountIfFormula(parameters);
                break;
            case 'averageif':
                formula = FormulaGenerator.generateAverageIfFormula(parameters);
                break;
            default:
                return res.status(400).json({
                    success: false,
                    error: `Unsupported conditional formula type: ${type}`
                });
        }

        res.json({
            success: true,
            formula,
            category: 'Conditional',
            validation: FormulaGenerator.validateFormula(formula.formula),
            relatedFormulas: FormulaGenerator.getRelatedFormulas(type, 'conditional'),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Validate Formula
router.post('/validate', async (req, res, next) => {
    try {
        const { formula } = req.body;
        
        if (!formula) {
            return res.status(400).json({
                success: false,
                error: 'Formula is required for validation'
            });
        }

        const validation = FormulaGenerator.validateFormula(formula);
        const extracted = FormulaGenerator.extractFormula(formula);

        res.json({
            success: true,
            validation,
            extracted,
            suggestions: validation.isValid ? [] : [
                'Check function syntax',
                'Verify range references',
                'Ensure parentheses are balanced',
                'Check for circular references'
            ],
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

// Formula Library
router.get('/library', async (req, res, next) => {
    try {
        const { category, search } = req.query;
        
        const library = {
            financial: [
                { name: 'NPV', description: 'Net Present Value', syntax: 'NPV(rate, value1, [value2], ...)' },
                { name: 'IRR', description: 'Internal Rate of Return', syntax: 'IRR(values, [guess])' },
                { name: 'PMT', description: 'Payment calculation', syntax: 'PMT(rate, nper, pv, [fv], [type])' },
                { name: 'FV', description: 'Future Value', syntax: 'FV(rate, nper, pmt, [pv], [type])' },
                { name: 'PV', description: 'Present Value', syntax: 'PV(rate, nper, pmt, [fv], [type])' }
            ],
            statistical: [
                { name: 'CORREL', description: 'Correlation coefficient', syntax: 'CORREL(array1, array2)' },
                { name: 'LINEST', description: 'Linear regression statistics', syntax: 'LINEST(known_y, [known_x], [const], [stats])' },
                { name: 'TTEST', description: 'T-test probability', syntax: 'TTEST(array1, array2, tails, type)' },
                { name: 'CONFIDENCE', description: 'Confidence interval', syntax: 'CONFIDENCE(alpha, standard_dev, size)' }
            ],
            lookup: [
                { name: 'VLOOKUP', description: 'Vertical lookup', syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])' },
                { name: 'XLOOKUP', description: 'Flexible lookup', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])' },
                { name: 'INDEX', description: 'Return value by position', syntax: 'INDEX(array, row_num, [column_num])' },
                { name: 'MATCH', description: 'Find position', syntax: 'MATCH(lookup_value, lookup_array, [match_type])' }
            ],
            conditional: [
                { name: 'IF', description: 'Conditional logic', syntax: 'IF(logical_test, value_if_true, value_if_false)' },
                { name: 'IFS', description: 'Multiple conditions', syntax: 'IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)' },
                { name: 'SUMIF', description: 'Conditional sum', syntax: 'SUMIF(range, criteria, [sum_range])' },
                { name: 'COUNTIF', description: 'Conditional count', syntax: 'COUNTIF(range, criteria)' }
            ]
        };

        let result = library;
        
        if (category && library[category]) {
            result = { [category]: library[category] };
        }
        
        if (search) {
            const searchLower = search.toLowerCase();
            result = Object.fromEntries(
                Object.entries(result).map(([cat, formulas]) => [
                    cat,
                    formulas.filter(f => 
                        f.name.toLowerCase().includes(searchLower) ||
                        f.description.toLowerCase().includes(searchLower)
                    )
                ])
            );
        }

        res.json({
            success: true,
            library: result,
            categories: Object.keys(library),
            totalFormulas: Object.values(library).reduce((sum, formulas) => sum + formulas.length, 0),
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        next(error);
    }
});

module.exports = router;
