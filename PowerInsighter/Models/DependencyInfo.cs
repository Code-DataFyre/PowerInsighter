namespace PowerInsighter.Models;

public class DependencyInfo
{
    public required string ObjectName { get; init; }
    public required string ObjectType { get; init; }
    public required string DependsOn { get; init; }
    public required string DependencyType { get; init; }
}
