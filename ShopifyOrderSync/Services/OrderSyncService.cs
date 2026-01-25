using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
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

        public event Action<string> LogEvent;
        public event Action<int> ProgressEvent;

        public OrderSyncService(
            string shopifyApiKey,
            string shopifyPassword,
            string shopifyShopDomain,
            string googleCredentialsJson,
            string spreadsheetId)
        {
            _shopifyApiKey = shopifyApiKey;
            _shopifyPassword = shopifyPassword;
            _shopifyShopDomain = shopifyShopDomain;
            _spreadsheetId = spreadsheetId;
            _httpClient = new HttpClient();
            _aiService = LocalAIService.Instance;

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

                    if (updates != null)
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

            var filter = new OrderListFilter
            {
                CreatedAtMin = startDate,
                CreatedAtMax = endDate,
                Status = "any"
            };

            var orders = await service.ListAsync(filter);
            return orders.Items.ToList();
        }

        private async Task<Dictionary<long, SheetOrderData>> GetExistingSheetDataAsync()
        {
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, "Sheet1!A2:F");
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
                            RowIndex = i + 2, // +2 because header is row 1, data starts at row 2
                            OrderId = orderId,
                            Status = row.Count > 3 ? row[3].ToString() : ""
                        };
                    }
                }
            }

            return existingOrders;
        }

        private OrderScenario DetermineScenario(Order order, Dictionary<long, SheetOrderData> existingOrders)
        {
            bool existsInSheet = existingOrders.ContainsKey(order.Id.Value);

            // Scenario B: Cancelled
            if (order.CancelledAt.HasValue)
            {
                return OrderScenario.Cancelled;
            }

            // Scenario A: New order
            if (!existsInSheet)
            {
                return OrderScenario.New;
            }

            var sheetData = existingOrders[order.Id.Value];

            // Scenario C: Skip if already delivered
            if (sheetData.Status.Equals("Delivered", StringComparison.OrdinalIgnoreCase))
            {
                return OrderScenario.Skip;
            }

            // Scenario D: Analyze tracking
            if (order.FulfillmentStatus == "fulfilled" && order.Fulfillments?.Any() == true)
            {
                return OrderScenario.Analyze;
            }

            // Scenario E: Stale unfulfilled
            if (order.FulfillmentStatus == "unfulfilled" &&
                order.CreatedAt.HasValue &&
                (DateTime.UtcNow - order.CreatedAt.Value).TotalDays > 3)
            {
                return OrderScenario.Stale;
            }

            return OrderScenario.Update;
        }

        private async Task<List<Request>> ProcessScenarioAsync(
            Order order,
            OrderScenario scenario,
            Dictionary<long, SheetOrderData> existingOrders)
        {
            var requests = new List<Request>();

            switch (scenario)
            {
                case OrderScenario.New:
                    Log($"  → Scenario A: New order - Appending to sheet");
                    requests.Add(CreateAppendRowRequest(order));
                    break;

                case OrderScenario.Cancelled:
                    Log($"  → Scenario B: Cancelled - Marking as cancelled");
                    if (existingOrders.ContainsKey(order.Id.Value))
                    {
                        int rowIndex = existingOrders[order.Id.Value].RowIndex;
                        requests.AddRange(CreateUpdateRequest(rowIndex, "Cancelled", "Grey"));
                    }
                    break;

                case OrderScenario.Skip:
                    Log($"  → Scenario C: Already delivered - Skipping");
                    break;

                case OrderScenario.Analyze:
                    Log($"  → Scenario D: Analyzing tracking status...");
                    var trackingUrl = order.Fulfillments.FirstOrDefault()?.TrackingUrl;

                    if (!string.IsNullOrEmpty(trackingUrl))
                    {
                        var analysisResult = await AnalyzeTrackingAsync(trackingUrl);
                        int rowIndex = existingOrders[order.Id.Value].RowIndex;
                        requests.AddRange(CreateUpdateRequest(rowIndex, analysisResult.Status, analysisResult.Color));
                        Log($"  → AI Analysis: {analysisResult.Status} ({analysisResult.Color})");
                    }
                    break;

                case OrderScenario.Stale:
                    Log($"  → Scenario E: Stale unfulfilled - Highlighting orange");
                    if (existingOrders.ContainsKey(order.Id.Value))
                    {
                        int rowIndex = existingOrders[order.Id.Value].RowIndex;
                        requests.AddRange(CreateUpdateRequest(rowIndex, "Unfulfilled - Stale", "Orange"));
                    }
                    break;
            }

            return requests;
        }

        private async Task<TrackingAnalysisResult> AnalyzeTrackingAsync(string trackingUrl)
        {
            try
            {
                Log($"    Downloading tracking page: {trackingUrl}");
                var htmlContent = await _httpClient.GetStringAsync(trackingUrl);

                Log($"    Analyzing with AI model...");
                var result = await _aiService.AnalyzeTrackingAsync(htmlContent);

                return result;
            }
            catch (Exception ex)
            {
                Log($"    Error analyzing tracking: {ex.Message}");
                return new TrackingAnalysisResult
                {
                    Status = "Analysis Failed",
                    Color = "Red",
                    ErrorMessage = ""
                };
            }
        }

        private Request CreateAppendRowRequest(Order order)
        {
            var trackingUrl = order.Fulfillments?.FirstOrDefault()?.TrackingUrl ?? "";

            return new Request
            {
                AppendCells = new AppendCellsRequest
                {
                    SheetId = 0,
                    Fields = "*",
                    Rows = new List<RowData>
                    {
                        new RowData
                        {
                            Values = new List<CellData>
                            {
                                new CellData { UserEnteredValue = new ExtendedValue { NumberValue = order.Id } },
                                new CellData { UserEnteredValue = new ExtendedValue { StringValue = order.Name } },
                                new CellData { UserEnteredValue = new ExtendedValue { StringValue = order.FinancialStatus } },
                                new CellData { UserEnteredValue = new ExtendedValue { StringValue = order.FulfillmentStatus ?? "unfulfilled" } },
                                new CellData { UserEnteredValue = new ExtendedValue { StringValue = trackingUrl } },
                                new CellData { UserEnteredValue = new ExtendedValue { StringValue = "New" } }
                            }
                        }
                    }
                }
            };
        }

        private List<Request> CreateUpdateRequest(int rowIndex, string status, string colorName)
        {
            var color = GetColor(colorName);

            return new List<Request>
            {
                // Update status text
                new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = 0,
                            StartRowIndex = rowIndex - 1,
                            EndRowIndex = rowIndex,
                            StartColumnIndex = 5,
                            EndColumnIndex = 6
                        },
                        Fields = "userEnteredValue",
                        Rows = new List<RowData>
                        {
                            new RowData
                            {
                                Values = new List<CellData>
                                {
                                    new CellData { UserEnteredValue = new ExtendedValue { StringValue = status } }
                                }
                            }
                        }
                    }
                },
                // Update background color
                new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = 0,
                            StartRowIndex = rowIndex - 1,
                            EndRowIndex = rowIndex,
                            StartColumnIndex = 0,
                            EndColumnIndex = 6
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = color
                            }
                        },
                        Fields = "userEnteredFormat.backgroundColor"
                    }
                }
            };
        }

        private Color GetColor(string colorName)
        {
            return colorName.ToLower() switch
            {
                "green" => new Color { Red = 0.7f, Green = 0.9f, Blue = 0.7f },
                "yellow" => new Color { Red = 1f, Green = 1f, Blue = 0.7f },
                "red" => new Color { Red = 0.95f, Green = 0.7f, Blue = 0.7f },
                "orange" => new Color { Red = 1f, Green = 0.85f, Blue = 0.6f },
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

    public enum OrderScenario
    {
        New,
        Cancelled,
        Skip,
        Analyze,
        Stale,
        Update
    }

    public class SheetOrderData
    {
        public int RowIndex { get; set; }
        public long OrderId { get; set; }
        public string Status { get; set; }
    }
}