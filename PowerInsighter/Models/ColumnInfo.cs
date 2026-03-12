namespace PowerInsighter.Models;

public class ColumnInfo
{
    public required string Name { get; init; }
    public required string Table { get; init; }
    public required string DataType { get; init; }
    public string? Expression { get; init; }
    public bool IsCalculated { get; init; }
    public bool IsHidden { get; init; }
    public string? Description { get; init; }
}
