using System.Collections.Concurrent;
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

        foreach (var rule in _rules)
        {
            if (string.IsNullOrWhiteSpace(rule.Expression) || string.IsNullOrWhiteSpace(rule.Scope))
                continue;

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
            }
            catch
            {
                // If a rule fails to compile or evaluate, skip it.
                // (Optional: surface an "invalid rule" diagnostics list separately later.)
            }
        }

        return violations;
    }

    private void Evaluate(BestPracticeRule rule, object target, string objectType, string objectName, string? tableName, List<BestPracticeViolation> violations)
    {
        var runner = GetOrCompile(rule.Scope!, rule.Expression!);

        var globals = new RuleGlobals(target);
        var isViolation = runner(globals).GetAwaiter().GetResult();

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
        // Allow TE-style expressions like: string.IsNullOrEmpty(Description)
        // We evaluate them in a context where "var This = (dynamic)Globals.This" and then expose members via dynamic.
        // This keeps most community rules working without requiring us to pre-map all scopes to types.
        return "\nusing System;\nusing System.Linq;\n\nvar it = (dynamic)Globals.This;\nreturn (" + expression + ") == true;\n";
    }

    private static string RewriteExpressionForDynamicTarget(string expression)
    {
        // Many BPA rules use unqualified member access (e.g. "Description", "IsHidden").
        // In Roslyn scripts those identifiers won't resolve unless we qualify them.
        // We do a best-effort rewrite to `it.<member>` while leaving already-qualified access intact.
        // This is intentionally conservative so more complex expressions still compile.

        var rewritten = expression;

        // If the rule already uses an explicit target (it./This./Model./Table.), leave it.
        if (rewritten.Contains("it.", StringComparison.OrdinalIgnoreCase) ||
            rewritten.Contains("This.", StringComparison.OrdinalIgnoreCase))
        {
            return rewritten;
        }

        // Common and safe replacements first.
        rewritten = Regex.Replace(rewritten, @"\bDescription\b", "it.Description");
        rewritten = Regex.Replace(rewritten, @"\bName\b", "it.Name");
        rewritten = Regex.Replace(rewritten, @"\bIsHidden\b", "it.IsHidden");
        rewritten = Regex.Replace(rewritten, @"\bDisplayFolder\b", "it.DisplayFolder");
        rewritten = Regex.Replace(rewritten, @"\bFormatString\b", "it.FormatString");
        rewritten = Regex.Replace(rewritten, @"\bExpression\b", "it.Expression");
        rewritten = Regex.Replace(rewritten, @"\bModifiedTime\b", "it.ModifiedTime");
        rewritten = Regex.Replace(rewritten, @"\bDataType\b", "it.DataType");
        rewritten = Regex.Replace(rewritten, @"\bDataCategory\b", "it.DataCategory");
        rewritten = Regex.Replace(rewritten, @"\bIsNullable\b", "it.IsNullable");
        rewritten = Regex.Replace(rewritten, @"\bIsKey\b", "it.IsKey");
        rewritten = Regex.Replace(rewritten, @"\bIsUnique\b", "it.IsUnique");

        // Avoid breaking explicitly-qualified accesses like `Table.Name` or `it.Description` by reverting over-qualification.
        rewritten = Regex.Replace(rewritten, @"\bit\.it\.", "it.");

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
