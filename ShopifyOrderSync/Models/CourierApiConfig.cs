namespace ShopifyOrderSync.Models
{
    public class CourierApiConfig
    {
        public string Name { get; set; } = string.Empty;
        public string DetectionUrl { get; set; } = string.Empty;
        public string ApiEndpoint { get; set; } = string.Empty;
        public List<string> QueryParameters { get; set; } = [];
        public bool Enabled { get; set; } = true;
    }
}
