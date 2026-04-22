using System.Data;
using System.Diagnostics;
using Microsoft.AnalysisServices.AdomdClient;
using PowerInsighter.Models;

namespace PowerInsighter.Services;

/// <summary>
/// Service for querying Dynamic Management Views (DMVs) from Analysis Services
/// to retrieve detailed model statistics and metadata.
/// </summary>
public class DmvQueryService
{
    /// <summary>
    /// Gets detailed storage statistics for all tables and columns in the model.
    /// Uses DMV queries: DISCOVER_STORAGE_TABLE_COLUMNS, DISCOVER_STORAGE_TABLES,
    /// and DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS.
    /// </summary>
    /// <param name="port">The port number of the Analysis Services instance</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Storage statistics including dictionary size, data size, and more</returns>
    public async Task<StorageStatistics?> GetStorageStatisticsAsync(int port, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            try
            {
                var connectionString = $"Data Source=localhost:{port};";
                using var connection = new AdomdConnection(connectionString);
                connection.Open();

                var stats = new StorageStatistics();

                // Query 1: Get column-level storage statistics (dictionary sizes)
                var columnStats = GetColumnStorageStatistics(connection);
                stats.TotalDictionarySize = columnStats.Sum(c => c.DictionarySize);
                stats.ColumnStatistics = columnStats;

                // Query 2: Get table-level storage statistics
                var tableStats = GetTableStorageStatistics(connection);
                stats.TableStatistics = tableStats;
                stats.TotalTableSize = tableStats.Sum(t => t.TotalSize);

                // Query 3: Get segment-level data sizes (actual data storage)
                var segmentStats = GetSegmentStorageStatistics(connection);
                stats.TotalDataSize = segmentStats.Sum(s => s.UsedSize);
                stats.TotalHierarchiesSize = segmentStats
                    .Where(s => s.ColumnType?.Contains("HIERARCHY", StringComparison.OrdinalIgnoreCase) == true)
                    .Sum(s => s.UsedSize);
                stats.TotalRelationshipSize = segmentStats
                    .Where(s => s.ColumnType?.Contains("RELATIONSHIP", StringComparison.OrdinalIgnoreCase) == true)
                    .Sum(s => s.UsedSize);

                connection.Close();

                Debug.WriteLine($"DMV: Storage statistics - Dictionary: {stats.TotalDictionarySize:N0} bytes, Data: {stats.TotalDataSize:N0} bytes, Hierarchies: {stats.TotalHierarchiesSize:N0} bytes, Relationships: {stats.TotalRelationshipSize:N0} bytes");
                return stats;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error querying DMV storage statistics: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }, cancellationToken);
    }

    /// <summary>
    /// Gets column-level storage statistics from DISCOVER_STORAGE_TABLE_COLUMNS DMV.
    /// This DMV provides DICTIONARY_SIZE per column.
    /// </summary>
    private List<ColumnStorageStatistics> GetColumnStorageStatistics(AdomdConnection connection)
    {
        var columnStats = new List<ColumnStorageStatistics>();

        try
        {
            // DISCOVER_STORAGE_TABLE_COLUMNS provides: TABLE_ID, COLUMN_ID, DICTIONARY_SIZE, COLUMN_ENCODING, COLUMN_TYPE
            var query = "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS";

            using var command = new AdomdCommand(query, connection);
            using var reader = command.ExecuteReader();

            // Log available columns for debugging
            var schemaTable = reader.GetSchemaTable();
            if (schemaTable != null)
            {
                var columnNames = string.Join(", ", schemaTable.Rows.Cast<DataRow>().Select(r => r["ColumnName"]?.ToString()));
                Debug.WriteLine($"DMV DISCOVER_STORAGE_TABLE_COLUMNS columns: {columnNames}");
            }

            while (reader.Read())
            {
                var stat = new ColumnStorageStatistics
                {
                    TableId = SafeGetString(reader, "TABLE_ID"),
                    ColumnId = SafeGetString(reader, "COLUMN_ID"),
                    DictionarySize = SafeGetLong(reader, "DICTIONARY_SIZE"),
                    ColumnType = SafeGetString(reader, "COLUMN_TYPE"),
                    ColumnEncoding = SafeGetString(reader, "COLUMN_ENCODING")
                };

                columnStats.Add(stat);
            }

            Debug.WriteLine($"DMV: Retrieved {columnStats.Count} column statistics, Total Dictionary: {columnStats.Sum(c => c.DictionarySize):N0} bytes");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error querying DISCOVER_STORAGE_TABLE_COLUMNS: {ex.Message}");
        }

        return columnStats;
    }

    /// <summary>
    /// Gets table-level storage statistics from DISCOVER_STORAGE_TABLES DMV.
    /// This DMV provides DICTIONARY_SIZE and USED_SIZE per table.
    /// </summary>
    private List<TableStorageStatistics> GetTableStorageStatistics(AdomdConnection connection)
    {
        var tableStats = new List<TableStorageStatistics>();

        try
        {
            // DISCOVER_STORAGE_TABLES provides: TABLE_ID, ROWS_COUNT, RIVIOLATION_COUNT, DICTIONARY_SIZE, USED_SIZE
            var query = "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES";

            using var command = new AdomdCommand(query, connection);
            using var reader = command.ExecuteReader();

            // Log available columns for debugging
            var schemaTable = reader.GetSchemaTable();
            if (schemaTable != null)
            {
                var columnNames = string.Join(", ", schemaTable.Rows.Cast<DataRow>().Select(r => r["ColumnName"]?.ToString()));
                Debug.WriteLine($"DMV DISCOVER_STORAGE_TABLES columns: {columnNames}");
            }

            while (reader.Read())
            {
                var stat = new TableStorageStatistics
                {
                    TableId = SafeGetString(reader, "TABLE_ID"),
                    RowsCount = SafeGetLong(reader, "ROWS_COUNT"),
                    DictionarySize = SafeGetLong(reader, "DICTIONARY_SIZE"),
                    UsedSize = SafeGetLong(reader, "USED_SIZE")
                };
                stat.TotalSize = stat.DictionarySize + stat.UsedSize;

                tableStats.Add(stat);
            }

            Debug.WriteLine($"DMV: Retrieved {tableStats.Count} table statistics, Total Used: {tableStats.Sum(t => t.UsedSize):N0} bytes, Total Dict: {tableStats.Sum(t => t.DictionarySize):N0} bytes");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error querying DISCOVER_STORAGE_TABLES: {ex.Message}");
        }

        return tableStats;
    }

    /// <summary>
    /// Gets segment-level storage statistics from DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS DMV.
    /// This provides the actual data sizes per segment including relationship and hierarchy columns.
    /// </summary>
    private List<SegmentStorageStatistics> GetSegmentStorageStatistics(AdomdConnection connection)
    {
        var segmentStats = new List<SegmentStorageStatistics>();

        try
        {
            var query = "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS";

            using var command = new AdomdCommand(query, connection);
            using var reader = command.ExecuteReader();

            // Log available columns for debugging
            var schemaTable = reader.GetSchemaTable();
            if (schemaTable != null)
            {
                var columnNames = string.Join(", ", schemaTable.Rows.Cast<DataRow>().Select(r => r["ColumnName"]?.ToString()));
                Debug.WriteLine($"DMV DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS columns: {columnNames}");
            }

            while (reader.Read())
            {
                var stat = new SegmentStorageStatistics
                {
                    TableId = SafeGetString(reader, "TABLE_ID"),
                    ColumnId = SafeGetString(reader, "COLUMN_ID"),
                    SegmentNumber = SafeGetLong(reader, "SEGMENT_NUMBER"),
                    UsedSize = SafeGetLong(reader, "USED_SIZE"),
                    ColumnType = SafeGetString(reader, "COLUMN_TYPE")
                };

                segmentStats.Add(stat);
            }

            Debug.WriteLine($"DMV: Retrieved {segmentStats.Count} segment statistics, Total Used: {segmentStats.Sum(s => s.UsedSize):N0} bytes");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error querying DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS: {ex.Message}");
        }

        return segmentStats;
    }

    /// <summary>
    /// Safely reads a string value from an AdomdDataReader by column name.
    /// Returns empty string if the column doesn't exist or value is null/DBNull.
    /// </summary>
    private string SafeGetString(AdomdDataReader reader, string columnName)
    {
        try
        {
            var ordinal = reader.GetOrdinal(columnName);
            if (reader.IsDBNull(ordinal))
                return string.Empty;
            return reader.GetValue(ordinal)?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Safely reads a long value from an AdomdDataReader by column name.
    /// Returns 0 if the column doesn't exist or value is null/DBNull.
    /// </summary>
    private long SafeGetLong(AdomdDataReader reader, string columnName)
    {
        try
        {
            var ordinal = reader.GetOrdinal(columnName);
            if (reader.IsDBNull(ordinal))
                return 0;

            var value = reader.GetValue(ordinal);
            if (value == null || value == DBNull.Value)
                return 0;

            if (long.TryParse(value.ToString(), out long result))
                return result;

            return 0;
        }
        catch
        {
            return 0;
        }
    }
}

/// <summary>
/// Contains comprehensive storage statistics from DMV queries.
/// </summary>
public class StorageStatistics
{
    public long TotalDictionarySize { get; set; }
    public long TotalDataSize { get; set; }
    public long TotalTableSize { get; set; }
    public long TotalRelationshipSize { get; set; }
    public long TotalHierarchiesSize { get; set; }
    public List<ColumnStorageStatistics> ColumnStatistics { get; set; } = [];
    public List<TableStorageStatistics> TableStatistics { get; set; } = [];

    public long TotalModelSize => TotalDictionarySize + TotalDataSize + TotalRelationshipSize + TotalHierarchiesSize;
}

/// <summary>
/// Storage statistics for a single column from DISCOVER_STORAGE_TABLE_COLUMNS.
/// </summary>
public class ColumnStorageStatistics
{
    public string TableId { get; set; } = string.Empty;
    public string ColumnId { get; set; } = string.Empty;
    public long DictionarySize { get; set; }
    public string ColumnType { get; set; } = string.Empty;
    public string ColumnEncoding { get; set; } = string.Empty;
}

/// <summary>
/// Storage statistics for a single table from DISCOVER_STORAGE_TABLES.
/// </summary>
public class TableStorageStatistics
{
    public string TableId { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public long RowsCount { get; set; }
    public long DictionarySize { get; set; }
    public long UsedSize { get; set; }
    public long TotalSize { get; set; }
}

/// <summary>
/// Storage statistics for a single segment from DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS.
/// </summary>
public class SegmentStorageStatistics
{
    public string TableId { get; set; } = string.Empty;
    public string ColumnId { get; set; } = string.Empty;
    public long SegmentNumber { get; set; }
    public long UsedSize { get; set; }
    public string ColumnType { get; set; } = string.Empty;
}
