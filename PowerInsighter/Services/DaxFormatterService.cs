using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PowerInsighter.Services;

public interface IDaxFormatterService
{
    Task<string> FormatDaxAsync(string dax);
}

public class DaxFormatterService : IDaxFormatterService
{
    private static readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(30)
    };

    private const string DaxFormatterUrl = "https://www.daxformatter.com/api/daxformatter/DaxFormat";

    public async Task<string> FormatDaxAsync(string dax)
    {
        if (string.IsNullOrWhiteSpace(dax))
            return dax;

        try
        {
            // DAX Formatter API expects form-urlencoded data
            var formData = new Dictionary<string, string>
            {
                { "fx", dax },
                { "s", "," },  // List separator (comma)
                { "d", "." }   // Decimal separator (period)
            };

            var content = new FormUrlEncodedContent(formData);

            var response = await _httpClient.PostAsync(DaxFormatterUrl, content);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            
            // The API returns the formatted DAX directly or a JSON with errors
            if (responseJson.StartsWith("{"))
            {
                var result = JsonSerializer.Deserialize<DaxFormatterResponse>(responseJson);
                
                if (result?.errors != null && result.errors.Count > 0)
                {
                    // Return original if there are errors
                    return dax;
                }
                
                return result?.formatted ?? dax;
            }
            
            // If it's not JSON, it's the formatted DAX directly
            return string.IsNullOrEmpty(responseJson) ? dax : responseJson;
        }
        catch (Exception)
        {
            // Return original DAX if formatting fails
            return dax;
        }
    }

    private class DaxFormatterResponse
    {
        [JsonPropertyName("formatted")]
        public string? formatted { get; set; }
        
        [JsonPropertyName("errors")]
        public List<DaxFormatterError>? errors { get; set; }
    }

    private class DaxFormatterError
    {
        [JsonPropertyName("line")]
        public int line { get; set; }
        
        [JsonPropertyName("column")]
        public int column { get; set; }
        
        [JsonPropertyName("message")]
        public string? message { get; set; }
    }
}
