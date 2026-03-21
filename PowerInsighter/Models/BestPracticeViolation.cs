namespace PowerInsighter.Models;

public sealed class BestPracticeViolation
{
    public required string RuleName { get; init; }
    public required string Category { get; init; }
    public required string Severity { get; init; }
    public required string ObjectType { get; init; }
    public required string ObjectName { get; init; }
    public string? TableName { get; init; }
    public required string Description { get; init; }
}
