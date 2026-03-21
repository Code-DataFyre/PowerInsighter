namespace PowerInsighter.Models;

public sealed class BestPracticeViolation
{
    public string RuleName { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Severity { get; init; } = string.Empty;
    public string ObjectType { get; init; } = string.Empty;
    public string ObjectName { get; init; } = string.Empty;
    public string? TableName { get; init; }
    public string Description { get; init; } = string.Empty;
}
