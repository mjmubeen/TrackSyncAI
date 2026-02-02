using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ShopifyOrderSync.Enums;
using ShopifyOrderSync.Models;
using ShopifySharp;
using ShopifySharp.Filters;
using System.Net.Http;

namespace ShopifyOrderSync.Services
{
    public class OrderSyncService
    {
        private readonly string _shopifyApiKey;
        private readonly string _shopifyPassword;
        private readonly string _shopifyShopDomain;
        private readonly string _spreadsheetId;
        private readonly SheetsService _sheetsService;
        private readonly HttpClient _httpClient;
        private readonly LocalAIService _aiService;
        private readonly List<CourierApiConfig> _courierAPIs;

        public event Action<string>? LogEvent;
        public event Action<int>? ProgressEvent;

        public OrderSyncService(
            string shopifyApiKey,
            string shopifyPassword,
            string shopifyShopDomain,
            string googleCredentialsJson,
            string spreadsheetId,
            List<CourierApiConfig> courierAPIs)
        {
            _shopifyApiKey = shopifyApiKey;
            _shopifyPassword = shopifyPassword;
            _shopifyShopDomain = shopifyShopDomain;
            _spreadsheetId = spreadsheetId;
            _httpClient = new HttpClient();
            _aiService = LocalAIService.Instance;
            _courierAPIs = courierAPIs ?? [];

            // Initialize Google Sheets
            var credential = GoogleCredential.FromJson(googleCredentialsJson)
                .CreateScoped(SheetsService.Scope.Spreadsheets);

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Shopify Order Sync"
            });
        }

        public async Task SyncOrdersAsync(DateTime startDate, DateTime endDate)
        {
            try
            {
                Log("=== Starting Order Sync ===");
                Log($"Date Range: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");

                // Fetch orders from Shopify
                var orders = await FetchShopifyOrdersAsync(startDate, endDate);
                Log($"Fetched {orders.Count} orders from Shopify");

                if (orders.Count == 0)
                {
                    Log("No orders found in date range.");
                    return;
                }

                // Get existing sheet data
                var existingOrders = await GetExistingSheetDataAsync();
                Log($"Found {existingOrders.Count} existing orders in Google Sheet");

                // Process each order
                int processedCount = 0;
                var batchUpdates = new List<Request>();

                foreach (var order in orders)
                {
                    processedCount++;
                    int progressPercent = (processedCount * 100) / orders.Count;
                    ProgressEvent?.Invoke(progressPercent);

                    Log($"\n[{processedCount}/{orders.Count}] Processing Order: {order.Name}");

                    var scenario = DetermineScenario(order, existingOrders);
                    var updates = await ProcessScenarioAsync(order, scenario, existingOrders);

                    if (updates != null && updates.Count > 0)
                    {
                        batchUpdates.AddRange(updates);
                    }

                    // Apply batch updates every 50 orders to avoid timeout
                    if (batchUpdates.Count >= 50)
                    {
                        await ApplyBatchUpdatesAsync(batchUpdates);
                        batchUpdates.Clear();
                    }
                }

                // Apply remaining updates
                if (batchUpdates.Count > 0)
                {
                    await ApplyBatchUpdatesAsync(batchUpdates);
                }

                Log("\n=== Sync Complete ===");
                Log($"Processed {processedCount} orders");
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.Message}");
                throw;
            }
        }

        private async Task<List<Order>> FetchShopifyOrdersAsync(DateTime startDate, DateTime endDate)
        {
            var service = new OrderService(_shopifyShopDomain, _shopifyPassword);
            var allOrders = new List<Order>();

            var filter = new OrderListFilter
            {
                CreatedAtMin = startDate,
                CreatedAtMax = endDate,
                Status = "any",
                Limit = 250  // Maximum allowed by Shopify API
            };

            Log("Fetching orders from Shopify (paginated)...");

            // Fetch first page
            var response = await service.ListAsync(filter);
            allOrders.AddRange(response.Items);
            Log($"  → Page 1: {response.Items.Count()} orders");

            // Check if there are more pages
            int pageCount = 1;
            while (response.HasNextPage)
            {
                pageCount++;
                Log($"  → Fetching page {pageCount}...");

                // Fetch next page using link info
                response = await service.ListAsync(response.GetNextPageFilter());
                allOrders.AddRange(response.Items);

                Log($"  → Page {pageCount}: {response.Items.Count()} orders (Total so far: {allOrders.Count})");

                // Safety limit to prevent infinite loops (adjust as needed)
                if (pageCount > 100)
                {
                    Log("  ⚠️ WARNING: Reached 100 pages limit. Stopping pagination.");
                    break;
                }
            }

            Log($"✓ Total orders fetched: {allOrders.Count} across {pageCount} page(s)");
            return allOrders;
        }

        private async Task<Dictionary<long, SheetOrderData>> GetExistingSheetDataAsync()
        {
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, "Sheet1!A2:L");
            var response = await request.ExecuteAsync();

            var existingOrders = new Dictionary<long, SheetOrderData>();

            if (response.Values != null)
            {
                for (int i = 0; i < response.Values.Count; i++)
                {
                    var row = response.Values[i];
                    if (row.Count > 0 && long.TryParse(row[0].ToString(), out long orderId))
                    {
                        existingOrders[orderId] = new SheetOrderData
                        {
                            RowIndex = i + 2,
                            OrderId = orderId,
                            CurrentStage = row.Count > 8 ? row[8]?.ToString() ?? "" : "",
                            WhatsAppStatus = row.Count > 9 ? row[9]?.ToString() ?? "" : "",
                            DeliveryStatus = row.Count > 10 ? row[10]?.ToString() ?? "" : "",
                            AIAlert = row.Count > 11 ? row[11]?.ToString() ?? "" : ""
                        };
                    }
                }
            }

            return existingOrders;
        }

        private static OrderScenario DetermineScenario(Order order, Dictionary<long, SheetOrderData> existingOrders)
        {
            var id = order.Id ?? 0;
            bool existsInSheet = existingOrders.ContainsKey(id);
            var tags = order.Tags ?? "";

            // Cancelled orders
            if (order.CancelledAt.HasValue || tags.Contains("Cancelled"))
            {
                return OrderScenario.Cancelled;
            }

            // New order - just placed
            if (!existsInSheet)
            {
                return OrderScenario.NewOrder;
            }

            var sheetData = existingOrders[id];

            // Stage 1: WhatsApp Confirmation
            if (tags.Contains("WhatsApp Sent") && !tags.Contains("Confirmed") && !tags.Contains("Did not pick up"))
            {
                return OrderScenario.AwaitingWhatsAppConfirm;
            }

            if (tags.Contains("Invalid WhatsApp"))
            {
                return OrderScenario.InvalidWhatsApp;
            }

            // Stage 2: Phone Verification
            if (tags.Contains("WhatsApp Confirmed") || tags.Contains("Awaiting Call"))
            {
                return OrderScenario.AwaitingPhoneCall;
            }

            if (tags.Contains("Did not pick up") || tags.Contains("No Answer"))
            {
                return OrderScenario.CustomerNotPickingPhone;
            }

            if (tags.Contains("Call Completed") && !tags.Contains("Size Confirmed"))
            {
                return OrderScenario.AwaitingSizeConfirmation;
            }

            // Stage 3: Ready for Courier
            if (tags.Contains("Size Confirmed") && order.FulfillmentStatus == "unfulfilled")
            {
                return OrderScenario.ReadyForCourier;
            }

            // Stage 4: Fulfilled - Track Delivery
            if (order.FulfillmentStatus == "fulfilled" && order.Fulfillments?.Any() == true)
            {
                // Skip if already delivered
                if (sheetData.DeliveryStatus.Equals("Delivered", StringComparison.OrdinalIgnoreCase))
                {
                    return OrderScenario.AlreadyDelivered;
                }

                return OrderScenario.TrackParcel;
            }

            // Stale orders (no progress in 24 hours)
            if (order.CreatedAt.HasValue && (DateTime.UtcNow - order.CreatedAt.Value).TotalHours > 24
                && order.FulfillmentStatus == "unfulfilled" && !tags.Contains("Size Confirmed"))
            {
                return OrderScenario.StaleOrder;
            }

            return OrderScenario.UpdateOnly;
        }

        private async Task<List<Request>> ProcessScenarioAsync(
            Order order,
            OrderScenario scenario,
            Dictionary<long, SheetOrderData> existingOrders)
        {
            var requests = new List<Request>();
            var id = order.Id ?? 0;

            switch (scenario)
            {
                case OrderScenario.NewOrder:
                    Log($"  → NEW ORDER: WhatsApp message sent");
                    requests.Add(CreateAppendRowRequest(order,
                        stage: "WhatsApp Sent",
                        whatsappStatus: "Pending",
                        deliveryStatus: "-",
                        aiAlert: "",
                        color: "LightBlue"));
                    break;

                case OrderScenario.AwaitingWhatsAppConfirm:
                    Log($"  → Awaiting WhatsApp confirmation");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value1))
                    {
                        requests.AddRange(CreateFullRowUpdate(value1.RowIndex, order,
                            stage: "Awaiting WhatsApp",
                            whatsappStatus: "Sent - No Response",
                            deliveryStatus: "-",
                            aiAlert: GetTimeBasedAlert(order, "WhatsApp not confirmed in"),
                            color: "Yellow"));
                    }
                    break;

                case OrderScenario.InvalidWhatsApp:
                    Log($"  → Invalid WhatsApp number - ALERT");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value2))
                    {
                        requests.AddRange(CreateFullRowUpdate(value2.RowIndex, order,
                            stage: "Invalid Contact",
                            whatsappStatus: "Invalid Number",
                            deliveryStatus: "-",
                            aiAlert: "⚠️ URGENT: Invalid WhatsApp - Manual contact needed",
                            color: "Red"));
                    }
                    break;

                case OrderScenario.AwaitingPhoneCall:
                    Log($"  → Ready for phone verification call");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value3))
                    {
                        requests.AddRange(CreateFullRowUpdate(value3.RowIndex, order,
                            stage: "Phone Verification",
                            whatsappStatus: "Confirmed",
                            deliveryStatus: "-",
                            aiAlert: "📞 Ready for call - Verify address & size",
                            color: "LightGreen"));
                    }
                    break;

                case OrderScenario.CustomerNotPickingPhone:
                    Log($"  → Customer not picking phone - ALERT");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value4))
                    {
                        requests.AddRange(CreateFullRowUpdate(value4.RowIndex, order,
                            stage: "Call Failed",
                            whatsappStatus: "Confirmed",
                            deliveryStatus: "-",
                            aiAlert: "⚠️ Customer not answering - Retry needed",
                            color: "Orange"));
                    }
                    break;

                case OrderScenario.AwaitingSizeConfirmation:
                    Log($"  → Call completed - awaiting size confirmation");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value5))
                    {
                        requests.AddRange(CreateFullRowUpdate(value5.RowIndex, order,
                            stage: "Size Confirmation",
                            whatsappStatus: "Confirmed",
                            deliveryStatus: "-",
                            aiAlert: "👟 Update shoe size in system",
                            color: "LightYellow"));
                    }
                    break;

                case OrderScenario.ReadyForCourier:
                    Log($"  → Ready for courier - Add to shipping");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value6))
                    {
                        requests.AddRange(CreateFullRowUpdate(value6.RowIndex, order,
                            stage: "Ready for Courier",
                            whatsappStatus: "Confirmed",
                            deliveryStatus: "Pending Pickup",
                            aiAlert: "📦 Add to courier system",
                            color: "Purple"));
                    }
                    break;

                case OrderScenario.TrackParcel:
                    Log($"  → Tracking parcel delivery with AI...");
                    var trackingUrl = order.Fulfillments?.FirstOrDefault()?.TrackingUrl;

                    if (!string.IsNullOrEmpty(trackingUrl))
                    {
                        var analysisResult = await AnalyzeTrackingAsync(trackingUrl);
                        var aiAlert = GenerateAIAlert(analysisResult, order);

                        requests.AddRange(CreateFullRowUpdate(existingOrders[id].RowIndex, order,
                            stage: "In Transit",
                            whatsappStatus: "Confirmed",
                            deliveryStatus: analysisResult.Status,
                            aiAlert: aiAlert,
                            color: analysisResult.Color));

                        Log($"  → AI Analysis: {analysisResult.Status} - {aiAlert}");
                    }
                    break;

                case OrderScenario.AlreadyDelivered:
                    Log($"  → Already delivered - Skipping");
                    break;

                case OrderScenario.StaleOrder:
                    Log($"  → STALE ORDER - No progress in 24h");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value7))
                    {
                        var currentStage = value7.CurrentStage;
                        requests.AddRange(CreateFullRowUpdate(value7.RowIndex, order,
                            stage: currentStage + " (STALE)",
                            whatsappStatus: value7.WhatsAppStatus,
                            deliveryStatus: value7.DeliveryStatus,
                            aiAlert: "⚠️ URGENT: Order stuck for 24+ hours - Investigate immediately",
                            color: "DarkRed"));
                    }
                    break;

                case OrderScenario.Cancelled:
                    Log($"  → Order cancelled");
                    if (existingOrders.TryGetValue(id, out SheetOrderData? value8))
                    {
                        requests.AddRange(CreateFullRowUpdate(value8.RowIndex, order,
                            stage: "Cancelled",
                            whatsappStatus: value8.WhatsAppStatus,
                            deliveryStatus: "Cancelled",
                            aiAlert: "",
                            color: "Grey"));
                    }
                    break;
            }

            return requests;
        }

        private async Task<TrackingAnalysisResult> AnalyzeTrackingAsync(string trackingUrl)
        {
            try
            {
                Log($"    Downloading tracking data: {trackingUrl}");

                // Try to use courier-specific API
                var htmlContent = await GetTrackingContentAsync(trackingUrl);

                if (string.IsNullOrWhiteSpace(htmlContent))
                {
                    Log($"    ⚠️ Warning: Empty response from tracking URL");
                    return new TrackingAnalysisResult
                    {
                        Status = "Tracking Unavailable",
                        Color = "Orange",
                        ErrorMessage = "Empty response from courier"
                    };
                }

                Log($"    → Raw content size: {htmlContent.Length} chars");

                // AI analysis
                var result = await _aiService.AnalyzeTrackingAsync(htmlContent);

                // Check if AI returned empty or invalid result
                if (string.IsNullOrWhiteSpace(result.Status) || result.Status == "Error")
                {
                    Log($"    ⚠️ AI analysis failed or returned empty result");
                    Log($"    → Status: '{result.Status}', Color: '{result.Color}'");
                    return new TrackingAnalysisResult
                    {
                        Status = "Analysis Failed",
                        Color = "Orange",
                        ErrorMessage = result.ErrorMessage ?? "AI returned empty response"
                    };
                }

                if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
                {
                    Log($"    ⚠️ AI analysis failed ErrorMessage='{result.ErrorMessage}'");
                }

                Log($"    ✓ AI Result: Status='{result.Status}', Color='{result.Color}'");
                return result;
            }
            catch (Exception ex)
            {
                Log($"    Error analyzing tracking: {ex.Message}");
                return new TrackingAnalysisResult
                {
                    Status = "Tracking Error",
                    Color = "Red",
                    ErrorMessage = ex.Message
                };
            }
        }

        private async Task<string> GetTrackingContentAsync(string trackingUrl)
        {
            // Check if any courier API matches
            foreach (var courierApi in _courierAPIs.Where(c => c.Enabled))
            {
                if (trackingUrl.Contains(courierApi.DetectionUrl, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        Log($"    → Detected {courierApi.Name} - Using API endpoint");

                        // Extract tracking ID from URL
                        var trackingId = ExtractTrackingId(trackingUrl, courierApi.QueryParameters);

                        if (!string.IsNullOrEmpty(trackingId))
                        {
                            // Replace {id} placeholder in API endpoint
                            string apiUrl = courierApi.ApiEndpoint.Replace("{id}", trackingId);
                            Log($"    → API URL: {apiUrl}");

                            var content = await _httpClient.GetStringAsync(apiUrl);

                            if (!string.IsNullOrWhiteSpace(content))
                            {
                                Log($"    ✓ Got API response ({content.Length} chars)");
                                return content;
                            }
                        }

                        Log($"    ⚠️ Could not extract tracking ID, falling back to direct URL");
                    }
                    catch (Exception ex)
                    {
                        Log($"    ⚠️ API call failed: {ex.Message}, trying direct URL");
                    }

                    // Break after first match to avoid multiple attempts
                    break;
                }
            }

            // Fallback: Regular HTML scraping
            Log($"    → Using direct URL scraping");
            return await _httpClient.GetStringAsync(trackingUrl);
        }

        private string? ExtractTrackingId(string trackingUrl, List<string> possibleParameters)
        {
            try
            {
                var uri = new Uri(trackingUrl);

                // Try query string parameters
                var query = System.Web.HttpUtility.ParseQueryString(uri.Query);

                foreach (var param in possibleParameters)
                {
                    var value = query[param];
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        Log($"    → Extracted ID from parameter '{param}': {value}");
                        return value;
                    }
                }

                // Try path segments (e.g., /track/12345678)
                var segments = uri.Segments;
                if (segments.Length > 0)
                {
                    var lastSegment = segments[^1].Trim('/');
                    if (!string.IsNullOrWhiteSpace(lastSegment) && lastSegment.Length > 5)
                    {
                        Log($"    → Extracted ID from path: {lastSegment}");
                        return lastSegment;
                    }
                }

                Log($"    ⚠️ Could not extract tracking ID from URL");
                return null;
            }
            catch (Exception ex)
            {
                Log($"    ⚠️ Error extracting tracking ID: {ex.Message}");
                return null;
            }
        }

        private static string GenerateAIAlert(TrackingAnalysisResult analysis, Order order)
        {
            var orderAge = order.CreatedAt.HasValue
                ? (DateTime.UtcNow - order.CreatedAt.Value).TotalDays
                : 0;

            return analysis.Status.ToLower() switch
            {
                "delivered" => "✅ Delivered successfully",
                "in-transit" or "in transit" => orderAge > 5
                    ? "⚠️ In transit > 5 days - Follow up with courier"
                    : "🚚 On the way",
                "stuck" => "🚨 URGENT: Parcel stuck - Contact courier immediately",
                "failed" => "🚨 CRITICAL: Delivery failed - Call customer NOW",
                "return" or "returned" => "⚠️ Parcel returned - Verify customer address",
                "customer not picking phone" => "📞 Courier can't reach customer - Call them",
                _ => $"ℹ️ Status: {analysis.Status}"
            };
        }

        private static string GetTimeBasedAlert(Order order, string message)
        {
            if (!order.CreatedAt.HasValue) return "";

            var hours = (DateTime.UtcNow - order.CreatedAt.Value).TotalHours;

            if (hours < 2) return "";
            if (hours < 6) return $"⏰ {message} {hours:F0} hours";
            if (hours < 24) return $"⚠️ {message} {hours:F0} hours - Follow up needed";
            return $"🚨 URGENT: {message} {hours / 24:F0} days";
        }

        private static Request CreateAppendRowRequest(Order order, string stage, string whatsappStatus,
            string deliveryStatus, string aiAlert, string color)
        {
            var trackingUrl = order.Fulfillments?.FirstOrDefault()?.TrackingUrl ?? "-";
            var bgColor = GetColor(color);

            // Extract customer name
            var customerName = order.Customer?.FirstName ?? "";
            if (!string.IsNullOrEmpty(order.Customer?.LastName))
            {
                customerName += " " + order.Customer.LastName;
            }
            if (string.IsNullOrWhiteSpace(customerName))
            {
                customerName = "Guest";
            }
            var phone = order.Customer?.Phone ?? "";
            if (string.IsNullOrWhiteSpace(phone))
            {
                var noteAttribute = order.NoteAttributes.FirstOrDefault(x=> x.Name == "Phone");
                if (noteAttribute != null)
                {
                    phone = noteAttribute.Value?.ToString() ?? "";
                }
                else
                {
                    var billingPhone = order.BillingAddress.Phone;
                    if (!string.IsNullOrEmpty(billingPhone))
                    {
                        phone = billingPhone;
                    }
                }
                    
            }
            // Format order date
            var orderDate = order.CreatedAt?.ToString("yyyy-MM-dd HH:mm") ?? "-";

            return new Request
            {
                AppendCells = new AppendCellsRequest
                {
                    SheetId = 0,
                    Fields = "*",
                    Rows =
                    [
                        new()
                        {
                            Values =
                            [
                                CreateCell(order.Id?.ToString() ?? "", bgColor),
                                CreateCell(order.Name ?? "", bgColor),
                                CreateCell(orderDate, bgColor),
                                CreateCell(customerName, bgColor),
                                CreateCell(phone, bgColor),
                                CreateCell(order.ShippingAddress?.City ?? "", bgColor),
                                CreateCell(order.FinancialStatus ?? "", bgColor),
                                CreateCell(trackingUrl, bgColor),
                                CreateCell(stage, bgColor),
                                CreateCell(whatsappStatus, bgColor),
                                CreateCell(deliveryStatus, bgColor),
                                CreateCell(aiAlert, bgColor)
                            ]
                        }
                    ]
                }
            };
        }

        private static List<Request> CreateFullRowUpdate(int rowIndex, Order order, string stage,
            string whatsappStatus, string deliveryStatus, string aiAlert, string color)
        {
            var trackingUrl = order.Fulfillments?.FirstOrDefault()?.TrackingUrl ?? "-";
            var bgColor = GetColor(color);

            // Extract customer name
            var customerName = order.Customer?.FirstName ?? "";
            if (!string.IsNullOrEmpty(order.Customer?.LastName))
            {
                customerName += " " + order.Customer.LastName;
            }
            if (string.IsNullOrWhiteSpace(customerName))
            {
                customerName = "Guest";
            }

            // Format order date
            var orderDate = order.CreatedAt?.ToString("yyyy-MM-dd HH:mm") ?? "-";

            return
            [
                new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = 0,
                            StartRowIndex = rowIndex - 1,
                            EndRowIndex = rowIndex,
                            StartColumnIndex = 0,
                            EndColumnIndex = 12
                        },
                        Fields = "*",
                        Rows =
                        [
                            new RowData
                            {
                                Values =
                                [
                                    CreateCell(order.Id?.ToString() ?? "", bgColor),
                                    CreateCell(order.Name ?? "", bgColor),
                                    CreateCell(orderDate, bgColor),
                                    CreateCell(customerName, bgColor),
                                    CreateCell(order.Customer?.Phone ?? "", bgColor),
                                    CreateCell(order.ShippingAddress?.City ?? "", bgColor),
                                    CreateCell(order.FinancialStatus ?? "", bgColor),
                                    CreateCell(trackingUrl, bgColor),
                                    CreateCell(stage, bgColor),
                                    CreateCell(whatsappStatus, bgColor),
                                    CreateCell(deliveryStatus, bgColor),
                                    CreateCell(aiAlert, bgColor)
                                ]
                            }
                        ]
                    }
                }
            ];
        }

        private static CellData CreateCell(string value, Color bgColor)
        {
            return new CellData
            {
                UserEnteredValue = new ExtendedValue { StringValue = value },
                UserEnteredFormat = new CellFormat { BackgroundColor = bgColor }
            };
        }

        private static Color GetColor(string colorName)
        {
            return colorName.ToLower() switch
            {
                "lightblue" => new Color { Red = 0.8f, Green = 0.9f, Blue = 1f },
                "lightgreen" => new Color { Red = 0.85f, Green = 0.95f, Blue = 0.85f },
                "lightyellow" => new Color { Red = 1f, Green = 1f, Blue = 0.85f },
                "yellow" => new Color { Red = 1f, Green = 1f, Blue = 0.7f },
                "orange" => new Color { Red = 1f, Green = 0.85f, Blue = 0.6f },
                "purple" => new Color { Red = 0.9f, Green = 0.8f, Blue = 1f },
                "green" => new Color { Red = 0.7f, Green = 0.9f, Blue = 0.7f },
                "red" => new Color { Red = 0.95f, Green = 0.7f, Blue = 0.7f },
                "darkred" => new Color { Red = 0.8f, Green = 0.4f, Blue = 0.4f },
                "grey" => new Color { Red = 0.85f, Green = 0.85f, Blue = 0.85f },
                _ => new Color { Red = 1f, Green = 1f, Blue = 1f }
            };
        }

        private async Task ApplyBatchUpdatesAsync(List<Request> requests)
        {
            if (requests.Count == 0) return;

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();
            Log($"  ✓ Applied {requests.Count} updates to Google Sheet");
        }

        private void Log(string message)
        {
            LogEvent?.Invoke(message);
        }
    }
}