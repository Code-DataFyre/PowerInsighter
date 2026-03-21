using System.Text.Json.Serialization;

namespace PowerInsighter.Models;

public sealed class BestPracticeRule
{
    [JsonPropertyName("ID")]
    public string? ID { get; init; }

    [JsonPropertyName("Name")]
    public string? Name { get; init; }

    [JsonPropertyName("Category")]
    public string? Category { get; init; }

    [JsonPropertyName("Description")]
    public string? Description { get; init; }

    [JsonPropertyName("Severity")]
    public int Severity { get; init; }

    [JsonPropertyName("Scope")]
    public string? Scope { get; init; }

    [JsonPropertyName("Expression")]
    public string? Expression { get; init; }

    [JsonIgnore]
    public string RuleId => ID ?? string.Empty;

    [JsonIgnore]
    public string RuleName => Name ?? string.Empty;

    [JsonIgnore]
    public string RuleCategory => Category ?? string.Empty;

    [JsonIgnore]
    public string RuleDescription => Description ?? string.Empty;
}
