namespace PowerInsighter.Models;

public class PowerBIInstance
{
    public required int Port { get; init; }
    public required int ProcessId { get; init; }
    public string? FileName { get; init; }
    public string? WindowTitle { get; init; }
    public DateTime? LastModified { get; init; }
    
    public string DisplayName => !string.IsNullOrEmpty(FileName) 
        ? $"{FileName} (Port: {Port})" 
        : !string.IsNullOrEmpty(WindowTitle)
            ? $"{WindowTitle} (Port: {Port})"
            : $"Power BI Desktop (Port: {Port})";
}
