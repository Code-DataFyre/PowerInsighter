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
}
