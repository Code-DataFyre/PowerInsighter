namespace PowerInsighter.Models;

public class RelationshipInfo
{
    public required string FromTable { get; init; }
    public required string FromColumn { get; init; }
    public required string ToTable { get; init; }
    public required string ToColumn { get; init; }
    public required string Cardinality { get; init; }
    public required string CrossFilterDirection { get; init; }
    public bool IsActive { get; init; }
}
