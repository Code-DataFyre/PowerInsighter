namespace PowerInsighter.Models;

public class ColumnInfo
{
    // Core Properties (Main fields)
    public required string Name { get; init; }
    public required string Table { get; init; }
    public required string DataType { get; init; }
    public string? Expression { get; init; }
    public bool IsCalculated { get; init; }
    public bool IsHidden { get; init; }
    public string? Description { get; init; }
    public string? DisplayFolder { get; init; }
    public string? FormatString { get; init; }
    
    // Advanced Properties (Settings fields)
    public string? SortByColumn { get; init; }
    public bool IsUnique { get; init; }
    public bool IsNullable { get; init; }
    public bool IsKey { get; init; }
    public string? SourceColumn { get; init; }
    public string? DataCategory { get; init; }
    public bool IsAvailableInMDX { get; init; }
    public string? State { get; init; }
    public string? ErrorMessage { get; init; }
    public DateTime? ModifiedTime { get; init; }
    public string? LineageTag { get; init; }
    public string? SummarizeBy { get; init; }
    public string? Encoding { get; init; }
}
