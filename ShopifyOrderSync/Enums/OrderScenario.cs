namespace ShopifyOrderSync.Enums
{
    public enum OrderScenario
    {
        NewOrder,
        AwaitingWhatsAppConfirm,
        InvalidWhatsApp,
        AwaitingPhoneCall,
        CustomerNotPickingPhone,
        AwaitingSizeConfirmation,
        ReadyForCourier,
        TrackParcel,
        AlreadyDelivered,
        StaleOrder,
        Cancelled,
        UpdateOnly
    }
}
