namespace PowerInsighter.Models;

/// <summary>
/// Detailed information about a table in the model, displayed in the Tables tab.
/// </summary>
public class TableDetailInfo
{
    public string Name { get; init; } = string.Empty;
    public int ColumnCount { get; init; }
    public int MeasureCount { get; init; }
    public int HierarchyCount { get; init; }
    public int CalculationGroupItemCount { get; init; }
    public bool IsCalculationGroup { get; init; }
    public string DataSource { get; init; } = string.Empty;
    public string SourceType { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public long RowCount { get; init; }
    public int PartitionCount { get; init; }
    public string? Description { get; init; }
}
