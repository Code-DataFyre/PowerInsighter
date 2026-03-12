namespace PowerInsighter.Models;

public class ModelOverview
{
    public required string ModelName { get; init; }
    public int TableCount { get; init; }
    public int MeasureCount { get; init; }
    public int ColumnCount { get; init; }
    public int RelationshipCount { get; init; }
    public int CalculatedColumnCount { get; init; }
    public int CalculatedTableCount { get; init; }
    public long ModelSize { get; init; }
    public DateTime? LastRefresh { get; init; }
    public string? CompatibilityLevel { get; init; }
}
