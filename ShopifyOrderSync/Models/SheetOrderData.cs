namespace ShopifyOrderSync.Models
{
    public class SheetOrderData
    {
        public int RowIndex { get; set; }
        public long OrderId { get; set; }
        public string CurrentStage { get; set; } = string.Empty;
        public string WhatsAppStatus { get; set; } = string.Empty;
        public string DeliveryStatus { get; set; } = string.Empty;
        public string AIAlert { get; set; } = string.Empty;
    }
}
