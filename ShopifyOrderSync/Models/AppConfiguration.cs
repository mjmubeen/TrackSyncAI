namespace ShopifyOrderSync.Models
{
    public class AppConfiguration
    {
        public required string ShopifyApiKey { get; set; } = string.Empty;
        public required string ShopifyPassword { get; set; } = string.Empty;
        public required string ShopifyShopDomain { get; set; } = string.Empty;
        public required GoogleCredentials GoogleCredentialsJson { get; set; }
        public required string SpreadsheetId { get; set; } = string.Empty;
        public List<CourierApiConfig> CourierAPIs { get; set; } = [];
    }
}
