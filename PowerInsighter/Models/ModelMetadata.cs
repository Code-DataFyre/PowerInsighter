namespace PowerInsighter.Models;

public class ModelMetadata
{
    public required string Name { get; init; }
    public required string Type { get; init; }
    public string? Parent { get; init; }
    public string? Details { get; init; }
}
