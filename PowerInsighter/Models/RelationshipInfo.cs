namespace PowerInsighter.Models;

public class RelationshipInfo
{
    // Core Properties (Main fields)
    public required string Name { get; init; }
    public required string FromTable { get; init; }
    public required string FromColumn { get; init; }
    public required string ToTable { get; init; }
    public required string ToColumn { get; init; }
    public required string Cardinality { get; init; }
    public required string CrossFilterDirection { get; init; }
    public bool IsActive { get; init; }
    
    // Additional Properties
    public string? FromCardinality { get; init; }
    public string? ToCardinality { get; init; }
    public string? SecurityFilteringBehavior { get; init; }
    public string? JoinOnDateBehavior { get; init; }
    public bool RelyOnReferentialIntegrity { get; init; }
    public string? State { get; init; }
    public DateTime? ModifiedTime { get; init; }
}
