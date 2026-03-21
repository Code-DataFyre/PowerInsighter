using System.Collections.Concurrent;
using System.Diagnostics;
using Microsoft.AnalysisServices.Tabular;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using PowerInsighter.Models;
using System.Text.RegularExpressions;

namespace PowerInsighter.Services;

public sealed class BestPracticeAnalyzer
{
    private readonly IReadOnlyList<BestPracticeRule> _rules;

    private readonly ConcurrentDictionary<string, ScriptRunner<bool>> _compiledRuleCache = new(StringComparer.Ordinal);

    public BestPracticeAnalyzer(IReadOnlyList<BestPracticeRule> rules)
    {
        _rules = rules;
    }

    public List<BestPracticeViolation> AnalyzeModel(Model model)
    {
        var violations = new List<BestPracticeViolation>();
        var rulesProcessed = 0;
        var rulesSkipped = 0;
        var rulesFailed = 0;

        Debug.WriteLine($"BPA: Starting analysis with {_rules.Count} rules");

        foreach (var rule in _rules)
        {
            if (string.IsNullOrWhiteSpace(rule.Expression) || string.IsNullOrWhiteSpace(rule.Scope))
            {
                rulesSkipped++;
                continue;
            }

            var scope = NormalizeScope(rule.Scope);

            try
            {
                switch (scope)
                {
                    case "model":
                        Evaluate(rule, model, objectType: "Model", objectName: model.Name, tableName: null, violations);
                        break;

                    case "table":
                        foreach (var table in model.Tables)
                            Evaluate(rule, table, objectType: "Table", objectName: table.Name, tableName: table.Name, violations);
                        break;

                    case "column":
                        foreach (var table in model.Tables)
                        {
                            foreach (var column in table.Columns)
                                Evaluate(rule, column, objectType: "Column", objectName: column.Name, tableName: table.Name, violations);
                        }
                        break;

                    case "measure":
                        foreach (var table in model.Tables)
                        {
                            foreach (var measure in table.Measures)
                                Evaluate(rule, measure, objectType: "Measure", objectName: measure.Name, tableName: table.Name, violations);
                        }
                        break;

                    case "relationship":
                        foreach (var rel in model.Relationships)
                        {
                            var name = rel.Name;
                            if (string.IsNullOrWhiteSpace(name))
                                name = $"{rel.FromTable?.Name} -> {rel.ToTable?.Name}";

                            Evaluate(rule, rel, objectType: "Relationship", objectName: name, tableName: null, violations);
                        }
                        break;
                }
                rulesProcessed++;
            }
            catch (Exception ex)
            {
                rulesFailed++;
                Debug.WriteLine($"BPA: Rule '{rule.RuleName}' failed: {ex.Message}");
            }
        }

        Debug.WriteLine($"BPA: Analysis complete. Processed: {rulesProcessed}, Skipped: {rulesSkipped}, Failed: {rulesFailed}, Violations: {violations.Count}");
        return violations;
    }

    private void Evaluate(BestPracticeRule rule, object target, string objectType, string objectName, string? tableName, List<BestPracticeViolation> violations)
    {
        try
        {
            bool isViolation;
            
            // Try SimpleRuleEvaluator first (faster, more reliable)
            if (SimpleRuleEvaluator.CanEvaluate(rule.Expression!))
            {
                isViolation = SimpleRuleEvaluator.Evaluate(rule.Expression!, target);
            }
            else
            {
                // Fall back to Roslyn for complex expressions
                var runner = GetOrCompile(rule.Scope!, rule.Expression!);
                var globals = new RuleGlobals(target);
                isViolation = runner(globals).GetAwaiter().GetResult();
            }

            if (!isViolation)
                return;

            violations.Add(new BestPracticeViolation
            {
                RuleName = rule.RuleName,
                Category = rule.RuleCategory,
                Severity = MapSeverity(rule.Severity),
                ObjectType = objectType,
                ObjectName = objectName,
                TableName = tableName,
                Description = rule.RuleDescription
            });

            Debug.WriteLine($"BPA: Violation found - {rule.RuleName} on {objectType} '{objectName}'");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"BPA: Evaluation error for rule '{rule.RuleName}' on {objectType} '{objectName}': {ex.Message}");
            throw;
        }
    }

    private ScriptRunner<bool> GetOrCompile(string scope, string expression)
    {
        var normalizedScope = NormalizeScope(scope);
        var rewrittenExpression = RewriteExpressionForDynamicTarget(expression);
        var cacheKey = $"{normalizedScope}::{rewrittenExpression}";

        return _compiledRuleCache.GetOrAdd(cacheKey, _ =>
        {
            var options = ScriptOptions.Default
                .WithReferences(
                    typeof(object).Assembly,
                    typeof(Enumerable).Assembly,
                    typeof(Model).Assembly)
                .WithImports(
                    "System",
                    "System.Linq",
                    "System.Text",
                    "Microsoft.AnalysisServices.Tabular");

            var scriptText = WrapExpression(rewrittenExpression);

            var script = CSharpScript.Create<bool>(scriptText, options, typeof(RuleGlobals));
            script.Compile();
            return script.CreateDelegate();
        });
    }

    private static string WrapExpression(string expression)
    {
        // Tabular Editor expressions have direct property access on the current object
        // We create a dynamic variable 'it' that references the object being evaluated
        // Then we evaluate the expression in that context
        return @"
using System;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices.Tabular;

dynamic it = This;
return (" + expression + @");
";
    }

    private static string RewriteExpressionForDynamicTarget(string expression)
    {
        // Many BPA rules use unqualified member access (e.g. "Description", "IsHidden").
        // In Roslyn scripts those identifiers won't resolve unless we qualify them with 'it.'
        // We do a best-effort rewrite to `it.<member>` while preserving Table., Model., and other qualified access

        var rewritten = expression;

        // If the rule already uses 'it.', leave it as is
        if (rewritten.Contains("it.", StringComparison.OrdinalIgnoreCase))
        {
            return rewritten;
        }

        // Common properties - replace only if not already qualified (not preceded by a dot)
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bDescription\b", "it.Description");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bName\b", "it.Name");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsHidden\b", "it.IsHidden");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bDisplayFolder\b", "it.DisplayFolder");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bFormatString\b", "it.FormatString");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bExpression\b", "it.Expression");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bModifiedTime\b", "it.ModifiedTime");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bDataType\b", "it.DataType");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bDataCategory\b", "it.DataCategory");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsNullable\b", "it.IsNullable");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsKey\b", "it.IsKey");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsUnique\b", "it.IsUnique");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsAvailableInMDX\b", "it.IsAvailableInMDX");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bSortByColumn\b", "it.SortByColumn");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bUsedInSortBy\b", "it.UsedInSortBy");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bUsedInHierarchies\b", "it.UsedInHierarchies");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bUsedInVariations\b", "it.UsedInVariations");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bUsedInRelationships\b", "it.UsedInRelationships");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bSummarizeBy\b", "it.SummarizeBy");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bObjectTypeName\b", "it.ObjectTypeName");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bPartitions\b", "it.Partitions");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bRowLevelSecurity\b", "it.RowLevelSecurity");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bCrossFilteringBehavior\b", "it.CrossFilteringBehavior");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bFromCardinality\b", "it.FromCardinality");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bToCardinality\b", "it.ToCardinality");
        rewritten = Regex.Replace(rewritten, @"(?<!\.)\bIsActive\b", "it.IsActive");

        return rewritten;
    }

    private static string NormalizeScope(string scope)
        => scope.Trim().ToLowerInvariant();

    private static string MapSeverity(int severity)
        => severity switch
        {
            <= 1 => "Information",
            2 => "Warning",
            _ => "Error"
        };

    private sealed class RuleGlobals
    {
        public RuleGlobals(object @this)
        {
            This = @this;
        }

        public object This { get; }
    }
}
