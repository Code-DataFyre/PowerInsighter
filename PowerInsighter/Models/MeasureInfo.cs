namespace PowerInsighter.Models;

public class MeasureInfo
{
    // Core Properties
    public required string Name { get; init; }
    public required string Table { get; init; }
    public string? Expression { get; init; }
    public string? Description { get; init; }
    public string? FormatString { get; init; }
    public bool IsHidden { get; init; }
    public string? DisplayFolder { get; init; }
    public string? DataType { get; init; }
    
    // Additional Properties
    public string? DetailRowsExpression { get; init; }
    public string? KPI { get; init; }
    public string? State { get; init; }
    public string? ErrorMessage { get; init; }
    public string? LineageTag { get; init; }
    
    // Metadata
    public DateTime? ModifiedTime { get; init; }
}
