using System.Net.Http.Headers;
using System.Net.Http;
using System.Text.Json;
using PowerInsighter.Models;

namespace PowerInsighter.Services;

public sealed class BestPracticeRulesRepository
{
    private static readonly Uri RulesApiUri = new("https://api.github.com/repos/microsoft/Analysis-Services/contents/BestPracticeRules?ref=master");

    private readonly HttpClient _httpClient;

    public BestPracticeRulesRepository(HttpClient httpClient)
    {
        _httpClient = httpClient;
        if (!_httpClient.DefaultRequestHeaders.UserAgent.Any())
        {
            _httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("PowerInsighter", "1.0"));
        }
    }

    public async Task<IReadOnlyList<BestPracticeRule>> LoadRulesAsync(CancellationToken cancellationToken = default)
    {
        using var response = await _httpClient.GetAsync(RulesApiUri, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);

        var rules = new List<BestPracticeRule>();

        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            if (!entry.TryGetProperty("name", out var nameProp))
                continue;

            var name = nameProp.GetString();
            if (string.IsNullOrWhiteSpace(name) || !name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                continue;

            if (!entry.TryGetProperty("download_url", out var downloadUrlProp))
                continue;

            var downloadUrl = downloadUrlProp.GetString();
            if (string.IsNullOrWhiteSpace(downloadUrl))
                continue;

            var fileRules = await LoadRulesFromFileAsync(new Uri(downloadUrl), cancellationToken).ConfigureAwait(false);
            rules.AddRange(fileRules);
        }

        return rules;
    }

    private async Task<IReadOnlyList<BestPracticeRule>> LoadRulesFromFileAsync(Uri downloadUri, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.GetAsync(downloadUri, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        // Most rule files are arrays of rule objects.
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

        var parsed = JsonSerializer.Deserialize<List<BestPracticeRule>>(json, options);
        return parsed ?? [];
    }
}
