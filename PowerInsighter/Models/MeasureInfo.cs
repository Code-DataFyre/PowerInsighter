namespace PowerInsighter.Models;

public class MeasureInfo
{
    public required string Name { get; init; }
    public required string Table { get; init; }
    public string? Expression { get; init; }
    public string? Description { get; init; }
    public string? FormatString { get; init; }
    public bool IsHidden { get; init; }
    public string? DisplayFolder { get; init; }
    public string? DataType { get; init; }
}
