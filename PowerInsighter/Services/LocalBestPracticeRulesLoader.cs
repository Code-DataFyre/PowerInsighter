using System.Diagnostics;
using System.IO;
using System.Text.Json;
using PowerInsighter.Models;

namespace PowerInsighter.Services;

/// <summary>
/// Loads Best Practice Rules from the local BPARules.json file
/// </summary>
public sealed class LocalBestPracticeRulesLoader
{
    private const string BPA_RULES_FILE = "ModelBestPractices/BPARules.json";

    public static async Task<List<BestPracticeRule>> LoadRulesAsync()
    {
        try
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, BPA_RULES_FILE);
            
            if (!File.Exists(filePath))
            {
                Debug.WriteLine($"BPA: Rules file not found at {filePath}");
                return new List<BestPracticeRule>();
            }

            Debug.WriteLine($"BPA: Loading rules from {filePath}");
            var json = await File.ReadAllTextAsync(filePath);

            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };

            var rules = JsonSerializer.Deserialize<List<BestPracticeRule>>(json, options) ?? new List<BestPracticeRule>();
            Debug.WriteLine($"BPA: Loaded {rules.Count} rules from local file");

            return rules;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"BPA: Error loading rules from file: {ex.Message}");
            return new List<BestPracticeRule>();
        }
    }
}
