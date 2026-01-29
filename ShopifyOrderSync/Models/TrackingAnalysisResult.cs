using Newtonsoft.Json;

namespace ShopifyOrderSync.Models
{
    /// <summary>
    /// Result of AI tracking analysis
    /// </summary>
    public class TrackingAnalysisResult
    {
        [JsonProperty("status")]
        public required string Status { get; set; }

        [JsonProperty("color")]
        public required string Color { get; set; }

        [JsonProperty("error")]
        public required string ErrorMessage { get; set; }
    }
}
