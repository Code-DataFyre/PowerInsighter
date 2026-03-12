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
                                    FileName = GetPowerBIFileName(parentProcess)
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
}
