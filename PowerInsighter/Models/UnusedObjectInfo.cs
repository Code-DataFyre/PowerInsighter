namespace PowerInsighter.Models;

public class UnusedObjectInfo
{
    public required string Name { get; init; }
    public required string ObjectType { get; init; }
    public required string Table { get; init; }
    public required string Reason { get; init; }
}
