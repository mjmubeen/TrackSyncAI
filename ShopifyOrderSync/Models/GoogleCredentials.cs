using System.Text.Json.Serialization;

namespace ShopifyOrderSync.Models
{
    public class GoogleCredentials
    {
        [JsonPropertyName("type")]
        public required string Type { get; set; }

        [JsonPropertyName("project_id")]
        public required string ProjectId { get; set; }

        [JsonPropertyName("private_key_id")]
        public required string PrivateKeyId { get; set; }

        [JsonPropertyName("private_key")]
        public required string PrivateKey { get; set; }

        [JsonPropertyName("client_email")]
        public required string ClientEmail { get; set; }

        [JsonPropertyName("client_id")]
        public required string ClientId { get; set; }

        [JsonPropertyName("auth_uri")]
        public required string AuthUri { get; set; }

        [JsonPropertyName("token_uri")]
        public required string TokenUri { get; set; }

        [JsonPropertyName("auth_provider_x509_cert_url")]
        public required string AuthProviderX509CertUrl { get; set; }

        [JsonPropertyName("client_x509_cert_url")]
        public required string ClientX509CertUrl { get; set; }

        [JsonPropertyName("universe_domain")]
        public required string UniverseDomain { get; set; }
    }
}
