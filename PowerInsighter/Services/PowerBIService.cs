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

    public async Task<List<int>> FindPowerBIPortsAsync(CancellationToken cancellationToken = default)
    {
        var foundPorts = new List<int>();

        // Method 1: Find msmdsrv.exe processes and their ports
        foundPorts.AddRange(await FindPortsFromProcessesAsync(cancellationToken));

        // Method 2: Try reading port files as fallback
        if (foundPorts.Count == 0)
        {
            Debug.WriteLine("Fallback: Trying port file method...");
            var portFromFile = GetPowerBIPortFromFile();
            if (portFromFile > 0)
            {
                foundPorts.Add(portFromFile);
            }
        }

        // Method 3: Scan common Power BI port range as last resort
        if (foundPorts.Count == 0)
        {
            Debug.WriteLine("Last resort: Scanning common port range...");
            foundPorts.AddRange(await ScanCommonPortsAsync(cancellationToken));
        }

        return foundPorts;
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

    private async Task<List<int>> FindPortsFromProcessesAsync(CancellationToken cancellationToken)
    {
        var foundPorts = new List<int>();

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
                            if (port > 0 && !foundPorts.Contains(port))
                            {
                                foundPorts.Add(port);
                                Debug.WriteLine($"? Detected port: {port}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error checking process: {ex.Message}");
                }
            }

            return foundPorts;
        }, cancellationToken);
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

    private static int GetPowerBIPortFromFile()
    {
        try
        {
            var basePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var workspacePath = Path.Combine(basePath, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces");

            if (!Directory.Exists(workspacePath))
                return 0;

            var workspaceFolders = Directory.GetDirectories(workspacePath)
                .OrderByDescending(f => new DirectoryInfo(f).LastWriteTime);

            foreach (var folder in workspaceFolders)
            {
                var portFile = Path.Combine(folder, "Data", "msmdsrv.port.txt");
                if (File.Exists(portFile))
                {
                    var content = File.ReadAllText(portFile, System.Text.Encoding.Unicode).Trim();
                    if (int.TryParse(content, out int port))
                        return port;
                }
            }
        }
        catch
        {
            // Swallow exception as this is a fallback mechanism
        }
        return 0;
    }

    private async Task<List<int>> ScanCommonPortsAsync(CancellationToken cancellationToken)
    {
        var foundPorts = new List<int>();

        var portsToScan = Enumerable.Range(PortScanStartRange, PortScanCount)
            .Where(p => p % PortScanInterval == 0)
            .ToList();

        Debug.WriteLine($"Scanning {portsToScan.Count} common ports...");

        foreach (var port in portsToScan)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (await IsPortListeningAsync(port, cancellationToken))
            {
                Debug.WriteLine($"Found listening port: {port}");
                foundPorts.Add(port);
                if (foundPorts.Count >= MaxScannedPorts)
                    break;
            }
        }

        return foundPorts;
    }

    private static async Task<bool> IsPortListeningAsync(int port, CancellationToken cancellationToken)
    {
        try
        {
            using var client = new TcpClient();
            var connectTask = client.ConnectAsync("localhost", port);
            var completedTask = await Task.WhenAny(connectTask, Task.Delay(PortCheckTimeoutMs, cancellationToken));

            if (completedTask == connectTask && !connectTask.IsFaulted)
            {
                return true;
            }
        }
        catch
        {
            // Port is not listening or connection failed
        }
        return false;
    }
}
