using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.AnalysisServices.Tabular;
using PowerInsighter.Models;

namespace PowerInsighter.Services;

/// <summary>
/// Simple rule evaluator that handles common BPA rule patterns without Roslyn compilation
/// This handles ~80% of rules that use simple property checks and comparisons
/// </summary>
public sealed class SimpleRuleEvaluator
{
    public static bool CanEvaluate(string expression)
    {
        // Check if this is a simple expression we can handle
        var normalized = expression.Replace("\r\n", " ").Replace("\n", " ").Trim();
        
        // Patterns we CAN handle:
        // - Simple equality: DataType = "Double"
        // - Simple boolean: IsHidden
        // - Negation: not IsHidden
        // - Contains: Expression.Contains("RELATED")
        // - StartsWith: Name.StartsWith("DateTableTemplate_")
        // - Method calls: string.IsNullOrEmpty(Description)
        // - Logical AND/OR: IsHidden and not UsedInSortBy.Any()
        
        // Patterns we CANNOT handle (need Roslyn):
        // - Complex LINQ: Model.AllMeasures.Any(...)
        // - Type conversions: Convert.ToInt64(...)
        // - Math operations: 1 / Math.Max(...)
        // - Nested object access: it.Table.Columns.Any(...)
        
        if (normalized.Contains("Model.") || 
            normalized.Contains("Convert.") || 
            normalized.Contains("Math.") ||
            normalized.Contains("AllMeasures") ||
            normalized.Contains("AllColumns") ||
            normalized.Contains("AllTables") ||
            normalized.Contains("GetAnnotation"))
        {
            return false; // Too complex, needs Roslyn
        }
        
        return true; // We can handle this
    }
    
    public static bool Evaluate(string expression, object target)
    {
        try
        {
            var normalized = expression.Replace("\r\n", " ").Replace("\n", " ").Trim();
            
            // Handle common patterns
            return EvaluateExpression(normalized, target);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Simple evaluator failed: {ex.Message}");
            return false;
        }
    }
    
    private static bool EvaluateExpression(string expr, object target)
    {
        // Handle logical operators (AND, OR)
        if (expr.Contains(" and ", StringComparison.OrdinalIgnoreCase))
        {
            var parts = SplitLogicalExpression(expr, "and");
            // If split failed (returned empty), the expression is too complex
            if (parts.Count == 0)
            {
                Debug.WriteLine($"Failed to split AND expression: {expr}");
                return false;
            }
            return parts.All(part => EvaluateExpression(part.Trim(), target));
        }
        
        if (expr.Contains(" or ", StringComparison.OrdinalIgnoreCase))
        {
            var parts = SplitLogicalExpression(expr, "or");
            // If split failed (returned empty), the expression is too complex
            if (parts.Count == 0)
            {
                Debug.WriteLine($"Failed to split OR expression: {expr}");
                return false;
            }
            return parts.Any(part => EvaluateExpression(part.Trim(), target));
        }
        
        // Handle negation
        if (expr.StartsWith("not ", StringComparison.OrdinalIgnoreCase))
        {
            var innerExpr = expr.Substring(4).Trim();
            return !EvaluateExpression(innerExpr, target);
        }
        
        // Handle parentheses
        if (expr.StartsWith("(") && expr.EndsWith(")"))
        {
            var inner = expr.Substring(1, expr.Length - 2).Trim();
            // Prevent infinite recursion if stripping parentheses doesn't change the expression
            if (inner.Length >= expr.Length - 2)
            {
                Debug.WriteLine($"Parentheses stripping failed: {expr}");
                return false;
            }
            return EvaluateExpression(inner, target);
        }
        
        // Handle equality comparisons
        if (expr.Contains(" = "))
        {
            return EvaluateEquality(expr, target);
        }
        
        // Handle inequality
        if (expr.Contains(" <> ") || expr.Contains(" != "))
        {
            var eqResult = EvaluateEquality(expr.Replace(" <> ", " = ").Replace(" != ", " = "), target);
            return !eqResult;
        }
        
        // Handle method calls
        if (expr.Contains("string.IsNullOrEmpty("))
        {
            return EvaluateStringIsNullOrEmpty(expr, target);
        }
        
        if (expr.Contains(".Contains("))
        {
            return EvaluateContains(expr, target);
        }
        
        if (expr.Contains(".StartsWith("))
        {
            return EvaluateStartsWith(expr, target);
        }
        
        if (expr.Contains(".EndsWith("))
        {
            return EvaluateEndsWith(expr, target);
        }
        
        if (expr.Contains(".Any("))
        {
            return EvaluateAny(expr, target);
        }
        
        // Handle simple boolean property
        return EvaluateBooleanProperty(expr, target);
    }
    
    private static List<string> SplitLogicalExpression(string expr, string op)
    {
        var parts = new List<string>();
        var depth = 0;
        var current = "";
        
        var pattern = $" {op} ";
        var i = 0;
        
        while (i < expr.Length)
        {
            var c = expr[i];
            
            if (c == '(') depth++;
            if (c == ')') depth--;
            
            if (depth == 0 && i <= expr.Length - pattern.Length)
            {
                var substring = expr.Substring(i, pattern.Length);
                if (substring.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                {
                    // Only split if we actually found the pattern
                    if (!string.IsNullOrWhiteSpace(current))
                    {
                        parts.Add(current.Trim());
                    }
                    current = "";
                    i += pattern.Length;
                    continue;
                }
            }
            
            current += c;
            i++;
        }
        
        if (!string.IsNullOrWhiteSpace(current))
        {
            parts.Add(current.Trim());
        }
        
        // If we didn't split anything, return empty list to avoid infinite recursion
        if (parts.Count == 1 && parts[0].Equals(expr.Trim(), StringComparison.Ordinal))
        {
            return new List<string>();
        }
        
        return parts;
    }
    
    private static bool EvaluateEquality(string expr, object target)
    {
        var parts = expr.Split(new[] { " = " }, StringSplitOptions.None);
        if (parts.Length != 2) return false;
        
        var left = parts[0].Trim();
        var right = parts[1].Trim().Trim('"');
        
        var leftValue = GetPropertyValue(left, target);
        
        if (leftValue == null)
            return right.Equals("null", StringComparison.OrdinalIgnoreCase);
        
        return leftValue.ToString()?.Equals(right, StringComparison.OrdinalIgnoreCase) == true;
    }
    
    private static bool EvaluateStringIsNullOrEmpty(string expr, object target)
    {
        var match = Regex.Match(expr, @"string\.IsNullOrEmpty\((\w+)\)", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        
        var propertyName = match.Groups[1].Value;
        var value = GetPropertyValue(propertyName, target);
        
        return string.IsNullOrEmpty(value?.ToString());
    }
    
    private static bool EvaluateContains(string expr, object target)
    {
        var match = Regex.Match(expr, @"(\w+)\.Contains\(""([^""]+)""\)", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        
        var propertyName = match.Groups[1].Value;
        var searchValue = match.Groups[2].Value;
        
        var value = GetPropertyValue(propertyName, target);
        
        return value?.ToString()?.Contains(searchValue, StringComparison.OrdinalIgnoreCase) == true;
    }
    
    private static bool EvaluateStartsWith(string expr, object target)
    {
        var match = Regex.Match(expr, @"(\w+)\.StartsWith\(""([^""]+)""\)", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        
        var propertyName = match.Groups[1].Value;
        var searchValue = match.Groups[2].Value;
        
        var value = GetPropertyValue(propertyName, target);
        
        return value?.ToString()?.StartsWith(searchValue, StringComparison.OrdinalIgnoreCase) == true;
    }
    
    private static bool EvaluateEndsWith(string expr, object target)
    {
        var match = Regex.Match(expr, @"(\w+)\.EndsWith\(""([^""]+)""\)", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        
        var propertyName = match.Groups[1].Value;
        var searchValue = match.Groups[2].Value;
        
        var value = GetPropertyValue(propertyName, target);
        
        return value?.ToString()?.EndsWith(searchValue, StringComparison.OrdinalIgnoreCase) == true;
    }
    
    private static bool EvaluateAny(string expr, object target)
    {
        var match = Regex.Match(expr, @"(\w+)\.Any\(\)", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        
        var propertyName = match.Groups[1].Value;
        var value = GetPropertyValue(propertyName, target);
        
        if (value is System.Collections.IEnumerable enumerable)
        {
            var enumerator = enumerable.GetEnumerator();
            return enumerator.MoveNext();
        }
        
        return false;
    }
    
    private static bool EvaluateBooleanProperty(string expr, object target)
    {
        // Simple property access
        var value = GetPropertyValue(expr.Trim(), target);
        
        if (value is bool boolValue)
            return boolValue;
        
        return false;
    }
    
    private static object? GetPropertyValue(string propertyPath, object target)
    {
        try
        {
            // Handle Table.PropertyName
            if (propertyPath.Contains("."))
            {
                var parts = propertyPath.Split('.');
                var current = target;
                
                foreach (var part in parts)
                {
                    if (current == null) return null;
                    
                    var prop = current.GetType().GetProperty(part);
                    if (prop == null) return null;
                    
                    current = prop.GetValue(current);
                }
                
                return current;
            }
            
            // Simple property
            var property = target.GetType().GetProperty(propertyPath);
            return property?.GetValue(target);
        }
        catch
        {
            return null;
        }
    }
}
