using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using Microsoft.AnalysisServices.Tabular;
using PowerInsighter.Models;

namespace PowerInsighter.Services;

public class PowerBIService : IPowerBIService
{
    private const string PowerBIProcessName = "PBIDesktop";
    private const string AnalysisServicesProcessName = "msmdsrv";
    private const int MinimumPort = 1024;
    private const int PortScanStartRange = 50000;
    private const int PortScanCount = 15000;
    private const int PortScanInterval = 500;
    private const int MaxScannedPorts = 5;
    private const int PortCheckTimeoutMs = 100;

    private readonly DmvQueryService _dmvQueryService;

    public PowerBIService()
    {
        _dmvQueryService = new DmvQueryService();
    }

    public bool IsPowerBIRunning()
    {
        try
        {
            var processes = Process.GetProcessesByName(PowerBIProcessName);
            return processes.Length > 0;
        }
        catch
        {
            return false;
        }
    }

    public async Task<List<PowerBIInstance>> FindPowerBIInstancesAsync(CancellationToken cancellationToken = default)
    {
        var instances = new List<PowerBIInstance>();

        return await Task.Run(() =>
        {
            var msmdsrvProcesses = Process.GetProcessesByName(AnalysisServicesProcessName);
            Debug.WriteLine($"Found {msmdsrvProcesses.Length} msmdsrv.exe process(es)");

            foreach (var process in msmdsrvProcesses)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    var parentId = GetParentProcessId(process.Id);
                    if (parentId > 0)
                    {
                        using var parentProcess = Process.GetProcessById(parentId);
                        if (parentProcess.ProcessName.Equals(PowerBIProcessName, StringComparison.OrdinalIgnoreCase))
                        {
                            Debug.WriteLine($"Found Power BI's msmdsrv.exe (PID: {process.Id})");

                            var port = FindPortForProcess(process.Id);
                            if (port > 0)
                            {
                                var instance = new PowerBIInstance
                                {
                                    Port = port,
                                    ProcessId = parentId,
                                    WindowTitle = GetProcessWindowTitle(parentProcess),
                                    FileName = GetPowerBIFileName(parentProcess),
                                    FilePath = GetPowerBIFilePath(parentProcess)
                                };

                                instances.Add(instance);
                                Debug.WriteLine($"? Created instance: {instance.DisplayName}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error checking process: {ex.Message}");
                }
            }

            // Fallback: Try reading port files
            if (instances.Count == 0)
            {
                Debug.WriteLine("Fallback: Trying port file method...");
                var portsFromFiles = GetPowerBIPortsFromFiles();
                foreach (var port in portsFromFiles)
                {
                    instances.Add(new PowerBIInstance
                    {
                        Port = port,
                        ProcessId = 0,
                        FileName = "Unknown (from port file)"
                    });
                }
            }

            return instances;
        }, cancellationToken);
    }

    public async Task<List<int>> FindPowerBIPortsAsync(CancellationToken cancellationToken = default)
    {
        var instances = await FindPowerBIInstancesAsync(cancellationToken);
        return instances.Select(i => i.Port).ToList();
    }

    public async Task<List<ModelMetadata>> LoadMetadataAsync(int port, CancellationToken cancellationToken = default)
    {
        var metadataList = new List<ModelMetadata>();

        await Task.Run(() =>
        {
            using var server = new Server();
            server.Connect($"DataSource=localhost:{port}");

            if (server.Databases.Count == 0)
            {
                throw new InvalidOperationException("No databases found.");
            }

            var model = server.Databases[0].Model;

            foreach (Table table in model.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();

                metadataList.Add(new ModelMetadata 
                { 
                    Name = table.Name, 
                    Type = "Table" 
                });

                foreach (Column col in table.Columns)
                {
                    metadataList.Add(new ModelMetadata 
                    { 
                        Name = col.Name, 
                        Type = "Column", 
                        Parent = table.Name 
                    });
                }
            }

            foreach (var rel in model.Relationships)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (rel is SingleColumnRelationship scr)
                {
                    metadataList.Add(new ModelMetadata
                    {
                        Name = "Relationship",
                        Type = "Link",
                        Details = $"{scr.FromTable.Name} -> {scr.ToTable.Name}"
                    });
                }
            }

            server.Disconnect();
        }, cancellationToken);

        return metadataList;
    }

    private static int GetParentProcessId(int processId)
    {
        try
        {
            var query = $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {processId}";
            using var searcher = new System.Management.ManagementObjectSearcher(query);
            foreach (var obj in searcher.Get())
            {
                return Convert.ToInt32(obj["ParentProcessId"]);
            }
        }
        catch
        {
            // Swallow exception as this is a fallback mechanism
        }
        return 0;
    }

    private static int FindPortForProcess(int processId)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "netstat",
                Arguments = "-ano",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process == null)
                return 0;

            var output = process.StandardOutput.ReadToEnd();
            var lines = output.Split('\n');

            foreach (var line in lines)
            {
                if (line.Contains("LISTENING") && line.Contains(processId.ToString()))
                {
                    var parts = line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        var address = parts[1];
                        if (address.Contains(':'))
                        {
                            var portStr = address.Split(':').Last();
                            if (int.TryParse(portStr, out int port) && port > MinimumPort)
                            {
                                return port;
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error finding port: {ex.Message}");
        }
        return 0;
    }

    private static string? GetProcessWindowTitle(Process process)
    {
        try
        {
            return !string.IsNullOrEmpty(process.MainWindowTitle) ? process.MainWindowTitle : null;
        }
        catch
        {
            return null;
        }
    }

    private static string? GetPowerBIFileName(Process process)
    {
        try
        {
            var windowTitle = process.MainWindowTitle;
            if (string.IsNullOrEmpty(windowTitle))
                return null;

            // Power BI Desktop window title format: "filename - Power BI Desktop"
            var parts = windowTitle.Split(new[] { " - Power BI Desktop" }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0)
            {
                var fileName = parts[0].Trim();
                // Remove asterisk if file is modified
                return fileName.TrimStart('*').Trim();
            }

            return windowTitle;
        }
        catch
        {
            return null;
        }
    }

    private static string? GetPowerBIFilePath(Process process)
    {
        try
        {
            // Try to get the file path from the process main module
            // Power BI Desktop opens the .pbix file, we need to search for recent files
            var fileName = GetPowerBIFileName(process);
            if (string.IsNullOrEmpty(fileName))
                return null;

            // Common locations where .pbix files might be
            var searchPaths = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            };

            // Search for the file in common locations
            foreach (var basePath in searchPaths)
            {
                if (Directory.Exists(basePath))
                {
                    var pbixFiles = Directory.GetFiles(basePath, "*.pbix", SearchOption.AllDirectories)
                        .Where(f => Path.GetFileNameWithoutExtension(f).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        .OrderByDescending(f => new FileInfo(f).LastWriteTime)
                        .FirstOrDefault();

                    if (pbixFiles != null)
                        return pbixFiles;
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error getting file path: {ex.Message}");
            return null;
        }
    }

    private static List<int> GetPowerBIPortsFromFiles()
    {
        var ports = new List<int>();
        try
        {
            var basePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var workspacePath = Path.Combine(basePath, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces");

            if (!Directory.Exists(workspacePath))
                return ports;

            var workspaceFolders = Directory.GetDirectories(workspacePath)
                .OrderByDescending(f => new DirectoryInfo(f).LastWriteTime);

            foreach (var folder in workspaceFolders)
            {
                var portFile = Path.Combine(folder, "Data", "msmdsrv.port.txt");
                if (File.Exists(portFile))
                {
                    var content = File.ReadAllText(portFile, System.Text.Encoding.Unicode).Trim();
                    if (int.TryParse(content, out int port))
                        ports.Add(port);
                }
            }
        }
        catch
        {
            // Swallow exception as this is a fallback mechanism
        }
        return ports;
    }

    public async Task<ModelOverview> GetModelOverviewAsync(int port, string? reportName = null, string? filePath = null, CancellationToken cancellationToken = default)
    {
        using var server = new Server();
        server.Connect($"DataSource=localhost:{port}");

        if (server.Databases.Count == 0)
            throw new InvalidOperationException("No databases found.");

        var database = server.Databases[0];
        var model = database.Model;

        var tableCount = model.Tables.Count;
        var measureCount = model.Tables.Sum(t => t.Measures.Count);
        var columnCount = model.Tables.Sum(t => t.Columns.Count);
        var relationshipCount = model.Relationships.Count;
        var calculatedColumnCount = model.Tables.Sum(t => t.Columns.Count(c => c.Type == ColumnType.Calculated));
        var calculatedTableCount = model.Tables.Count(t => t.Partitions.Any(p => p.SourceType == PartitionSourceType.Calculated));

        // Get .pbix file size from disk
        long fileSize = 0;
        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                fileSize = fileInfo.Length;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not get file size: {ex.Message}");
            }
        }

        // Additional statistics
        var hiddenTableCount = model.Tables.Count(t => t.IsHidden);
        var hiddenMeasureCount = model.Tables.Sum(t => t.Measures.Count(m => m.IsHidden));
        var hiddenColumnCount = model.Tables.Sum(t => t.Columns.Count(c => c.IsHidden));
        var partitionCount = model.Tables.Sum(t => t.Partitions.Count);

        // Calculation Groups (available in compatibility level 1500+)
        var calculationGroupCount = 0;
        try
        {
            calculationGroupCount = model.Tables.Count(t => t.CalculationGroup != null);
        }
        catch
        {
            // CalculationGroup property may not be available in older models
        }

        // Get accurate storage statistics from DMV queries
        long totalDictionarySize = 0;
        long totalDataSize = 0;
        long totalHierarchiesSize = 0;

        try
        {
            Debug.WriteLine("Querying DMV for accurate storage statistics...");
            var storageStats = await _dmvQueryService.GetStorageStatisticsAsync(port, cancellationToken);

            if (storageStats != null)
            {
                totalDictionarySize = storageStats.TotalDictionarySize;
                totalDataSize = storageStats.TotalDataSize;
                totalHierarchiesSize = storageStats.TotalHierarchiesSize;

                Debug.WriteLine($"DMV Statistics - Dictionary: {totalDictionarySize:N0} bytes, Data: {totalDataSize:N0} bytes, Hierarchies: {totalHierarchiesSize:N0} bytes");
            }
            else
            {
                // Fallback to estimated size if DMV query fails
                Debug.WriteLine("DMV query failed, using fallback estimated size");
                totalDictionarySize = database.EstimatedSize;
                totalDataSize = database.EstimatedSize;

                var hierarchyCount = model.Tables.Sum(t => t.Hierarchies.Count);
                totalHierarchiesSize = hierarchyCount * 1024; // 1KB per hierarchy estimate
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error getting storage statistics: {ex.Message}");
            // Fallback to estimated size
            totalDictionarySize = database.EstimatedSize;
            totalDataSize = database.EstimatedSize;
            var hierarchyCount = model.Tables.Sum(t => t.Hierarchies.Count);
            totalHierarchiesSize = hierarchyCount * 1024;
        }

        // Use reportName if provided, otherwise fall back to model.Name or database.Name
        var modelName = !string.IsNullOrEmpty(reportName) ? reportName : (model.Name ?? database.Name);

            var overview = new ModelOverview
            {
                ModelName = modelName,
                TableCount = tableCount,
                MeasureCount = measureCount,
                ColumnCount = columnCount,
                RelationshipCount = relationshipCount,
                CalculatedColumnCount = calculatedColumnCount,
                CalculatedTableCount = calculatedTableCount,
                ModelSize = database.EstimatedSize,
                LastRefresh = database.LastUpdate,
                CompatibilityLevel = database.CompatibilityLevel.ToString(),
                HiddenTableCount = hiddenTableCount,
                HiddenMeasureCount = hiddenMeasureCount,
                HiddenColumnCount = hiddenColumnCount,
                PartitionCount = partitionCount,
                DefaultMode = model.DefaultMode.ToString(),
                Culture = model.Culture,
                CreatedTimestamp = database.CreatedTimestamp,
                StructureModifiedTime = model.ModifiedTime,
                CalculationGroupCount = calculationGroupCount,
                TotalDictionarySize = totalDictionarySize,
                TotalDataSize = totalDataSize,
                TotalHierarchiesSize = totalHierarchiesSize,
                FileSize = fileSize,
                FilePath = filePath
            };

        server.Disconnect();
        return overview;
    }

    public async Task<List<MeasureInfo>> GetMeasuresAsync(int port, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var measures = new List<MeasureInfo>();

            using var server = new Server();
            server.Connect($"DataSource=localhost:{port}");

            if (server.Databases.Count == 0)
                throw new InvalidOperationException("No databases found.");

            var model = server.Databases[0].Model;

            foreach (Table table in model.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();

                foreach (Measure measure in table.Measures)
                {
                    // Get DetailRowsExpression safely (may not exist in older versions)
                    string? detailRowsExpr = null;
                    try
                    {
                        var detailRowsProp = measure.GetType().GetProperty("DetailRowsExpression");
                        if (detailRowsProp?.GetValue(measure) is object detailRows)
                        {
                            var exprProp = detailRows.GetType().GetProperty("Expression");
                            detailRowsExpr = exprProp?.GetValue(detailRows)?.ToString();
                        }
                    }
                    catch { /* Property may not exist */ }

                    measures.Add(new MeasureInfo
                    {
                        Name = measure.Name,
                        Table = table.Name,
                        Expression = measure.Expression,
                        Description = measure.Description,
                        FormatString = measure.FormatString,
                        IsHidden = measure.IsHidden,
                        DisplayFolder = measure.DisplayFolder,
                        DataType = measure.DataType.ToString(),
                        DetailRowsExpression = detailRowsExpr,
                        KPI = measure.KPI?.TargetExpression,
                        State = measure.State.ToString(),
                        ErrorMessage = measure.ErrorMessage,
                        LineageTag = measure.LineageTag,
                        ModifiedTime = measure.ModifiedTime
                    });
                }
            }

            server.Disconnect();
            return measures;
        }, cancellationToken);
    }

    public async Task<List<ColumnInfo>> GetColumnsAsync(int port, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var columns = new List<ColumnInfo>();

            using var server = new Server();
            server.Connect($"DataSource=localhost:{port}");

            if (server.Databases.Count == 0)
                throw new InvalidOperationException("No databases found.");

            var model = server.Databases[0].Model;

            foreach (Table table in model.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();

                foreach (Column column in table.Columns)
                {
                    var isCalculated = column.Type == ColumnType.Calculated;
                    string? expression = null;
                    
                    if (isCalculated && column is CalculatedColumn calcCol)
                    {
                        expression = calcCol.Expression;
                    }

                    // Get SortByColumn name if set
                    string? sortByColumnName = null;
                    try
                    {
                        sortByColumnName = column.SortByColumn?.Name;
                    }
                    catch { /* May not be available */ }

                    // Get SourceColumn for DataColumns
                    string? sourceColumn = null;
                    if (column is DataColumn dataCol)
                    {
                        sourceColumn = dataCol.SourceColumn;
                    }

                    columns.Add(new ColumnInfo
                    {
                        Name = column.Name,
                        Table = table.Name,
                        DataType = column.DataType.ToString(),
                        IsCalculated = isCalculated,
                        Expression = expression,
                        IsHidden = column.IsHidden,
                        Description = column.Description,
                        DisplayFolder = column.DisplayFolder,
                        FormatString = column.FormatString,
                        SortByColumn = sortByColumnName,
                        IsUnique = column.IsUnique,
                        IsNullable = column.IsNullable,
                        IsKey = column.IsKey,
                        SourceColumn = sourceColumn,
                        DataCategory = column.DataCategory,
                        IsAvailableInMDX = column.IsAvailableInMDX,
                        State = column.State.ToString(),
                        ErrorMessage = column.ErrorMessage,
                        ModifiedTime = column.ModifiedTime,
                        LineageTag = column.LineageTag,
                        SummarizeBy = column.SummarizeBy.ToString(),
                        Encoding = column.EncodingHint.ToString()
                    });
                }
            }

            server.Disconnect();
            return columns;
        }, cancellationToken);
    }

    public async Task<List<RelationshipInfo>> GetRelationshipsAsync(int port, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var relationships = new List<RelationshipInfo>();

            using var server = new Server();
            server.Connect($"DataSource=localhost:{port}");

            if (server.Databases.Count == 0)
                throw new InvalidOperationException("No databases found.");

            var model = server.Databases[0].Model;

            foreach (var relationship in model.Relationships)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (relationship is SingleColumnRelationship scr)
                {
                    relationships.Add(new RelationshipInfo
                    {
                        Name = scr.Name,
                        FromTable = scr.FromTable.Name,
                        FromColumn = scr.FromColumn.Name,
                        ToTable = scr.ToTable.Name,
                        ToColumn = scr.ToColumn.Name,
                        Cardinality = scr.FromCardinality.ToString() + " to " + scr.ToCardinality.ToString(),
                        FromCardinality = scr.FromCardinality.ToString(),
                        ToCardinality = scr.ToCardinality.ToString(),
                        CrossFilterDirection = scr.CrossFilteringBehavior.ToString(),
                        SecurityFilteringBehavior = scr.SecurityFilteringBehavior.ToString(),
                        JoinOnDateBehavior = scr.JoinOnDateBehavior.ToString(),
                        RelyOnReferentialIntegrity = scr.RelyOnReferentialIntegrity,
                        IsActive = scr.IsActive,
                        State = scr.State.ToString(),
                        ModifiedTime = scr.ModifiedTime
                    });
                }
            }

            server.Disconnect();
            return relationships;
        }, cancellationToken);
    }

    public async Task<List<UnusedObjectInfo>> GetUnusedObjectsAsync(int port, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var unusedObjects = new List<UnusedObjectInfo>();

            using var server = new Server();
            server.Connect($"DataSource=localhost:{port}");

            if (server.Databases.Count == 0)
                throw new InvalidOperationException("No databases found.");

            var model = server.Databases[0].Model;

            var allMeasures = model.Tables
                .SelectMany(t => t.Measures.Select(m => (Table: t.Name, Measure: m)))
                .ToList();

            // Precompute all measure expressions for basic reference checks
            var expressions = allMeasures
                .Select(m => m.Measure.Expression)
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .ToList();

            bool IsMeasureReferenced(string measureName)
                => expressions.Any(e => e.Contains($"[{measureName}]", StringComparison.OrdinalIgnoreCase));

            bool IsColumnReferenced(string tableName, string columnName)
            {
                var token = $"'{tableName}'[{columnName}]";
                return expressions.Any(e => e.Contains(token, StringComparison.OrdinalIgnoreCase));
            }

            // Unused Measures
            foreach (var (tableName, measure) in allMeasures)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var referenced = IsMeasureReferenced(measure.Name);

                if (measure.IsHidden && !referenced)
                {
                    unusedObjects.Add(new UnusedObjectInfo
                    {
                        Name = measure.Name,
                        ObjectType = "Measure",
                        Table = tableName,
                        Reason = "Hidden measure with no references"
                    });
                }
                else if (!referenced)
                {
                    unusedObjects.Add(new UnusedObjectInfo
                    {
                        Name = measure.Name,
                        ObjectType = "Measure",
                        Table = tableName,
                        Reason = "Measure not referenced by other measures"
                    });
                }
            }

            // Unused Columns (only meaningful for hidden columns here)
            foreach (Table table in model.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();

                foreach (Column column in table.Columns)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Heuristic: flag hidden columns that are never referenced in any measure
                    if (!column.IsHidden)
                        continue;

                    var referenced = IsColumnReferenced(table.Name, column.Name);
                    if (!referenced)
                    {
                        unusedObjects.Add(new UnusedObjectInfo
                        {
                            Name = column.Name,
                            ObjectType = "Column",
                            Table = table.Name,
                            Reason = "Hidden column with no references in measure expressions"
                        });
                    }
                }
            }

            server.Disconnect();
            return unusedObjects;
        }, cancellationToken);
    }

    /// <summary>
    /// Gets detailed storage statistics from DMV queries.
    /// This provides accurate per-column dictionary size, data size, and more.
    /// </summary>
    /// <param name="port">The port number of the Analysis Services instance</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Detailed storage statistics or null if query fails</returns>
    public async Task<StorageStatistics?> GetStorageStatisticsAsync(int port, CancellationToken cancellationToken = default)
    {
        try
        {
            return await _dmvQueryService.GetStorageStatisticsAsync(port, cancellationToken);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error in GetStorageStatisticsAsync: {ex.Message}");
            return null;
        }
    }
}
