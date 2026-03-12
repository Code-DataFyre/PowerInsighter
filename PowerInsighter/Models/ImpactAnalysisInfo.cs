namespace PowerInsighter.Models;

public class ImpactAnalysisInfo
{
    public required string ObjectName { get; init; }
    public required string ObjectType { get; init; }
    public required string ImpactedObject { get; init; }
    public required string ImpactType { get; init; }
    public required string Severity { get; init; }
}
