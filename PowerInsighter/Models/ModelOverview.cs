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

    // Additional useful properties
    public int HiddenTableCount { get; init; }
    public int HiddenMeasureCount { get; init; }
    public int HiddenColumnCount { get; init; }
    public int PartitionCount { get; init; }
    public string? DefaultMode { get; init; }
    public string? Culture { get; init; }
    public DateTime? CreatedTimestamp { get; init; }
    public DateTime? StructureModifiedTime { get; init; }

    // New properties
    public int CalculationGroupCount { get; init; }
    public long TotalDictionarySize { get; init; }
    public long TotalDataSize { get; init; }
    public long TotalHierarchiesSize { get; init; }

    // File size properties
    public long FileSize { get; init; }  // .pbix file size on disk
    public string? FilePath { get; init; }  // Path to the .pbix file

    // Data source properties
    public List<DataSourceInfo> DataSources { get; init; } = [];
    public List<TableSourceInfo> TableSources { get; init; } = [];
    public int DataSourceCount { get; init; }
    public int UniqueSourceTypesCount { get; init; }  // Total unique source types used
}

/// <summary>
/// Information about a data source defined in the model.
/// Available directly from model.DataSources collection.
/// </summary>
public class DataSourceInfo
{
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;  // Provider or Structured
    public string? Description { get; init; }
    public string? ConnectionDetails { get; init; }
    public int MaxConnections { get; init; }
    public DateTime? ModifiedTime { get; init; }
}

/// <summary>
/// Information about a table's data source type.
/// Requires iterating through tables and partitions.
/// </summary>
public class TableSourceInfo
{
    public string TableName { get; init; } = string.Empty;
    public string SourceType { get; init; } = string.Empty;  // M, Query, Calculated, None, etc.
    public string? SourceExpression { get; init; }  // M expression or query text
    public string? DetectedSourceKind { get; init; }  // Parsed from M: Sql.Database, Excel.Workbook, etc.
}
