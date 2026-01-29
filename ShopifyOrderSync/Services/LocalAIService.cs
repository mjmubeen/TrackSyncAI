using LLama;
using LLama.Common;
using LLama.Sampling;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ShopifyOrderSync.Enums;
using ShopifyOrderSync.Models;

namespace ShopifyOrderSync.Services
{
    /// <summary>
    /// Singleton service for managing the local AI model using LLamaSharp
    /// </summary>
    public class LocalAIService
    {
        private static LocalAIService? _instance;
        private static readonly object _lock = new();

        private LLamaWeights? _model;
        private LLamaContext? _context;
        private InteractiveExecutor? _executor;
        private bool _isLoaded = false;
        private long _memoryUsageBytes = 0;

        public bool IsLoaded => _isLoaded;
        public long MemoryUsageMB => _memoryUsageBytes / (1024 * 1024);

        private LocalAIService() { }

        public static LocalAIService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new LocalAIService();
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Initialize the AI model from a .gguf file
        /// </summary>
        public async Task<bool> LoadModelAsync(string modelPath, Action<string>? logCallback = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    logCallback?.Invoke($"Loading model from: {modelPath}");

                    // Dispose existing model if loaded
                    UnloadModel();

                    // Model parameters
                    var parameters = new ModelParams(modelPath)
                    {
                        ContextSize = 4096,      // Context window
                        GpuLayerCount = 0,       // Use CPU only (set to > 0 for GPU)
                        Embeddings = false
                    };

                    logCallback?.Invoke("Initializing LLamaWeights...");
                    _model = LLamaWeights.LoadFromFile(parameters);

                    logCallback?.Invoke("Creating inference context...");
                    _context = _model.CreateContext(parameters);

                    logCallback?.Invoke("Creating interactive executor...");
                    _executor = new InteractiveExecutor(_context);

                    // Estimate memory usage
                    _memoryUsageBytes = GC.GetTotalMemory(false);

                    _isLoaded = true;
                    logCallback?.Invoke($"✓ Model loaded successfully! Memory: {MemoryUsageMB} MB");

                    return true;
                }
                catch (Exception ex)
                {
                    logCallback?.Invoke($"✗ Error loading model: {ex.Message}");
                    _isLoaded = false;
                    return false;
                }
            });
        }

        /// <summary>
        /// Analyze tracking content using the AI model
        /// </summary>
        public async Task<TrackingAnalysisResult> AnalyzeTrackingAsync(string trackingContent)
        {
            if (!_isLoaded)
            {
                throw new InvalidOperationException("Model not loaded. Please load the model first.");
            }

            if (_executor == null)
            {
                return new TrackingAnalysisResult
                {
                    Status = "Error",
                    Color = "Red",
                    ErrorMessage = "AI executor is not initialized."
                };
            }

            return await Task.Run(async () =>
            {
                try
                {
                    // Detect content type and prepare accordingly
                    var contentType = DetectContentType(trackingContent);
                    string preparedText = PrepareContentByType(trackingContent, contentType);

                    // Safety check for empty content
                    if (string.IsNullOrWhiteSpace(preparedText))
                    {
                        return new TrackingAnalysisResult
                        {
                            Status = "In-Transit",
                            Color = "Yellow",
                            ErrorMessage = "Empty tracking content after processing"
                        };
                    }

                    // Construct optimized prompt based on content type
                    string prompt = BuildPrompt(preparedText, contentType);

                    var inferenceParams = new InferenceParams
                    {
                        SamplingPipeline = new DefaultSamplingPipeline()
                        {
                            Temperature = 0.2f,  // Low temperature for consistent JSON output
                            Seed = 1337
                        },
                        AntiPrompts = ["<|end|>", "<|user|>", "\n\n"],
                        MaxTokens = 100
                    };

                    // Get AI response with early stopping
                    string response = "";
                    int tokenCount = 0;
                    await foreach (var token in _executor.InferAsync(prompt, inferenceParams))
                    {
                        response += token;
                        tokenCount++;

                        // Early stop if we have valid JSON
                        if (response.Contains('}') && response.Contains("status") && response.Contains("color"))
                        {
                            break;
                        }

                        // Safety: Stop if response is too long
                        if (tokenCount > 150)
                        {
                            break;
                        }
                    }

                    // Check if response is empty
                    if (string.IsNullOrWhiteSpace(response))
                    {
                        return new TrackingAnalysisResult
                        {
                            Status = "In-Transit",
                            Color = "Yellow",
                            ErrorMessage = "AI returned empty response - defaulting to In-Transit"
                        };
                    }

                    // Parse and validate response
                    return ParseAIResponse(response);
                }
                catch (Exception ex)
                {
                    return new TrackingAnalysisResult
                    {
                        Status = "Error",
                        Color = "Red",
                        ErrorMessage = ex.Message
                    };
                }
            });
        }

        /// <summary>
        /// Detect what type of content we're dealing with
        /// </summary>
        private static ContentType DetectContentType(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
                return ContentType.Unknown;

            var trimmed = content.TrimStart();

            // Check for JSON
            if (trimmed.StartsWith('{') || trimmed.StartsWith('['))
            {
                try
                {
                    JToken.Parse(trimmed);
                    return ContentType.JSON;
                }
                catch
                {
                    // Not valid JSON, continue checking
                }
            }

            // Check for XML
            if (trimmed.StartsWith("<?xml") || trimmed.StartsWith('<'))
            {
                if (trimmed.Contains("</") && trimmed.Contains('>'))
                {
                    return ContentType.XML;
                }
            }

            // Check for HTML
            if (trimmed.Contains("<html") || trimmed.Contains("<body") ||
                trimmed.Contains("<div") || trimmed.Contains("<script"))
            {
                return ContentType.HTML;
            }

            // Default to plain text
            return ContentType.PlainText;
        }

        /// <summary>
        /// Prepare content based on detected type
        /// </summary>
        private static string PrepareContentByType(string content, ContentType type)
        {
            return type switch
            {
                ContentType.JSON => ProcessJSON(content),
                ContentType.XML => ProcessXML(content),
                ContentType.HTML => ProcessHTML(content),
                ContentType.PlainText => ProcessPlainText(content),
                _ => content
            };
        }

        /// <summary>
        /// Process JSON API response - Extract only tracking-relevant fields
        /// </summary>
        private static string ProcessJSON(string json)
        {
            try
            {
                var jToken = JToken.Parse(json);
                var extracted = new System.Text.StringBuilder();

                // Priority 1: Critical tracking fields (most important)
                var criticalFields = new[]
                {
                    "status", "delivery_status", "tracking_status", "shipment_status",
                    "current_status", "order_status", "state", "stage", "step",
                    "delivered", "is_delivered", "deliveryStatus"
                };

                // Priority 2: Location and time fields
                var locationFields = new[]
                {
                    "location", "current_location", "last_location", "city",
                    "destination", "origin", "hub", "facility"
                };

                var timeFields = new[]
                {
                    "date", "timestamp", "updated_at", "delivery_date",
                    "expected_delivery", "estimated_delivery", "delivered_at"
                };

                // Priority 3: Additional context
                var contextFields = new[]
                {
                    "remarks", "message", "description", "details", "notes",
                    "reason", "comment", "failed_reason", "exception"
                };

                // Extract in priority order
                ExtractFields(jToken, criticalFields, extracted, "STATUS");
                ExtractFields(jToken, locationFields, extracted, "LOCATION");
                ExtractFields(jToken, timeFields, extracted, "TIME");
                ExtractFields(jToken, contextFields, extracted, "DETAILS");

                // Handle tracking history/events (array of events)
                ExtractTrackingHistory(jToken, extracted);

                var result = extracted.ToString();

                // If extraction yielded nothing useful, return cleaned JSON
                if (string.IsNullOrWhiteSpace(result) || result.Length < 20)
                {
                    return CleanJSON(json);
                }

                // Limit to 1500 chars (most important info should be at top)
                return result.Length > 1500 ? result[..1500] : result;
            }
            catch
            {
                // If JSON parsing fails, treat as plain text
                return ProcessPlainText(json);
            }
        }

        /// <summary>
        /// Extract fields from JSON token
        /// </summary>
        private static void ExtractFields(JToken token, string[] fieldNames, System.Text.StringBuilder output, string category)
        {
            var found = new HashSet<string>();

            foreach (var fieldName in fieldNames)
            {
                var values = token.SelectTokens($"..{fieldName}").ToList();

                foreach (var value in values)
                {
                    if (value?.Type == JTokenType.String || value?.Type == JTokenType.Boolean)
                    {
                        var strValue = value.ToString();
                        if (!string.IsNullOrWhiteSpace(strValue) && !found.Contains(strValue))
                        {
                            output.AppendLine($"[{category}] {fieldName}: {strValue}");
                            found.Add(strValue);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Extract tracking history/events from JSON
        /// </summary>
        private static void ExtractTrackingHistory(JToken token, System.Text.StringBuilder output)
        {
            // Common field names for tracking events
            var historyFields = new[] { "history", "events", "tracking_history", "shipment_history", "timeline", "updates" };

            foreach (var field in historyFields)
            {
                var events = token.SelectTokens($"..{field}").FirstOrDefault();

                if (events?.Type == JTokenType.Array)
                {
                    var eventArray = (JArray)events;

                    // Get last 3 events (most recent)
                    var recentEvents = eventArray.Reverse().Take(3).Reverse();

                    output.AppendLine("\n[RECENT EVENTS]");
                    foreach (var evt in recentEvents)
                    {
                        var status = evt["status"]?.ToString() ?? evt["message"]?.ToString() ?? evt["description"]?.ToString();
                        var date = evt["date"]?.ToString() ?? evt["timestamp"]?.ToString() ?? evt["time"]?.ToString();
                        var location = evt["location"]?.ToString() ?? evt["city"]?.ToString();

                        if (!string.IsNullOrWhiteSpace(status))
                        {
                            output.Append($"- {status}");
                            if (!string.IsNullOrWhiteSpace(date)) output.Append($" ({date})");
                            if (!string.IsNullOrWhiteSpace(location)) output.Append($" at {location}");
                            output.AppendLine();
                        }
                    }

                    break; // Only process first matching history field
                }
            }
        }

        /// <summary>
        /// Clean JSON by removing unnecessary whitespace
        /// </summary>
        private static string CleanJSON(string json)
        {
            // Remove extra whitespace while preserving structure
            var cleaned = System.Text.RegularExpressions.Regex.Replace(json, @"\s+", " ");
            return cleaned.Length > 1500 ? cleaned[..1500] : cleaned;
        }

        /// <summary>
        /// Process XML response - Extract tracking-relevant nodes
        /// </summary>
        private static string ProcessXML(string xml)
        {
            try
            {
                // Extract text content between tags
                var pattern = @"<(?:status|location|date|time|message|description|remarks|delivery)[^>]*>([^<]+)</";
                var matches = System.Text.RegularExpressions.Regex.Matches(xml, pattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                var sb = new System.Text.StringBuilder();
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        var value = System.Net.WebUtility.HtmlDecode(match.Groups[1].Value.Trim());
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            sb.AppendLine(value);
                        }
                    }
                }

                var result = sb.ToString();
                return string.IsNullOrWhiteSpace(result) ? ProcessPlainText(xml) : result;
            }
            catch
            {
                return ProcessPlainText(xml);
            }
        }

        /// <summary>
        /// Process HTML page - Remove scripts, styles, extract text
        /// </summary>
        private static string ProcessHTML(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return "";

            // Remove script and style tags with content
            html = System.Text.RegularExpressions.Regex.Replace(
                html, @"<script[^>]*>.*?</script>", "",
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            html = System.Text.RegularExpressions.Regex.Replace(
                html, @"<style[^>]*>.*?</style>", "",
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Remove HTML comments
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<!--.*?-->", "");

            // Remove all HTML tags
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<[^>]+>", " ");

            // Decode HTML entities
            html = System.Net.WebUtility.HtmlDecode(html);

            // Remove extra whitespace
            html = System.Text.RegularExpressions.Regex.Replace(html, @"\s+", " ").Trim();

            // Apply intelligent truncation
            return TruncateIntelligently(html, 1800);
        }

        /// <summary>
        /// Process plain text - Keep as is with intelligent truncation
        /// </summary>
        private static string ProcessPlainText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            // Remove excessive whitespace
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();

            // Apply intelligent truncation
            return TruncateIntelligently(text, 1800);
        }

        /// <summary>
        /// Intelligent truncation - Keeps important parts based on keywords
        /// </summary>
        private static string TruncateIntelligently(string text, int maxLength)
        {
            if (text.Length <= maxLength)
                return text;

            // Keywords that indicate important tracking information
            var keywords = new[]
            {
                "delivered", "delivery", "status", "tracking",
                "failed", "returned", "return", "stuck", "transit",
                "location", "date", "time", "received", "recipient",
                "out for delivery", "in transit", "picked up",
                "attempted", "exception", "delay", "completed"
            };

            // Split into sentences or lines
            var segments = text.Split(['.', '\n', ';'], StringSplitOptions.RemoveEmptyEntries);
            var importantSegments = new List<string>();
            var currentLength = 0;

            // First pass: Collect segments with keywords
            foreach (var segment in segments)
            {
                var lowerSegment = segment.ToLower();
                if (keywords.Any(k => lowerSegment.Contains(k)))
                {
                    if (currentLength + segment.Length < maxLength)
                    {
                        importantSegments.Add(segment.Trim());
                        currentLength += segment.Length + 2; // +2 for spacing
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // If we found enough important segments, use those
            if (importantSegments.Count > 0 && currentLength > maxLength / 2)
            {
                return string.Join(". ", importantSegments);
            }

            // Otherwise, take beginning and end
            var halfLength = maxLength / 2 - 10;
            var beginning = text[..Math.Min(halfLength, text.Length)];
            var endStart = Math.Max(halfLength, text.Length - halfLength);
            var end = text[endStart..];

            return beginning + " [...] " + end;
        }

        /// <summary>
        /// Build optimized prompt based on content type
        /// </summary>
        private static string BuildPrompt(string preparedText, ContentType type)
        {
            var contentDescription = type switch
            {
                ContentType.JSON => "JSON API response with tracking fields",
                ContentType.XML => "XML tracking data",
                ContentType.HTML => "web tracking page content",
                _ => "tracking information"
            };

            return $@"<|system|>
You are a courier tracking analyzer for an e-commerce business. Analyze {contentDescription} and determine delivery status.
Return ONLY valid JSON in this EXACT format: {{""status"": ""Status"", ""color"": ""Color""}}

Status options (pick ONE):
- Delivered: Package successfully delivered to customer
- In-Transit: Moving normally through courier network
- Stuck: No movement for 2+ days OR shows delay/exception
- Failed: Delivery attempt failed
- Return: Being returned to sender
- Customer Not Picking Phone: Courier cannot contact customer

Color codes (pick ONE):
- Green: Delivered
- Yellow: In-Transit (normal)
- Orange: Stuck (warning)
- Red: Failed, Return, or Customer unreachable

Keywords to look for:
✓ DELIVERED: delivered, received, completed, successful delivery
✓ IN-TRANSIT: in transit, on the way, out for delivery, picked up, moving
✓ STUCK: delay, exception, stuck, no movement, hold
✓ FAILED: failed, unsuccessful, not delivered, cancelled
✓ RETURN: returned, return to sender, rto
<|end|>
<|user|>
Analyze this {contentDescription} and return ONLY JSON:

{preparedText}

Return JSON now:<|end|>
<|assistant|>";
        }

        /// <summary>
        /// Parse AI response to extract JSON
        /// </summary>
        private static TrackingAnalysisResult ParseAIResponse(string aiResponse)
        {
            try
            {
                // Try to find JSON in response
                int jsonStart = aiResponse.IndexOf('{');
                int jsonEnd = aiResponse.LastIndexOf('}');

                if (jsonStart >= 0 && jsonEnd > jsonStart)
                {
                    string json = aiResponse.Substring(jsonStart, jsonEnd - jsonStart + 1);
                    var result = JsonConvert.DeserializeObject<TrackingAnalysisResult>(json)
                        ?? throw new Exception("Deserialization returned null");

                    // Validate and normalize result
                    result.Status = NormalizeStatus(result.Status);
                    result.Color = NormalizeColor(result.Color);

                    return result;
                }

                // Fallback: Try to parse entire response as JSON
                var fallbackResult = JsonConvert.DeserializeObject<TrackingAnalysisResult>(aiResponse)
                    ?? throw new Exception("Deserialization returned null");

                fallbackResult.Status = NormalizeStatus(fallbackResult.Status);
                fallbackResult.Color = NormalizeColor(fallbackResult.Color);

                return fallbackResult;
            }
            catch
            {
                // Ultimate fallback
                return new TrackingAnalysisResult
                {
                    Status = "In-Transit",
                    Color = "Yellow",
                    ErrorMessage = "Could not parse AI response"
                };
            }
        }

        /// <summary>
        /// Normalize status to valid values
        /// </summary>
        private static string NormalizeStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return "In-Transit";

            status = status.Trim();
            var lower = status.ToLower();

            // Map variations to standard statuses
            if (lower.Contains("deliver") && !lower.Contains("not") && !lower.Contains("fail"))
                return "Delivered";
            if (lower.Contains("transit") || lower.Contains("on the way") || lower.Contains("moving"))
                return "In-Transit";
            if (lower.Contains("stuck") || lower.Contains("delay") || lower.Contains("hold"))
                return "Stuck";
            if (lower.Contains("fail") || lower.Contains("unsuccess") || lower.Contains("cancel"))
                return "Failed";
            if (lower.Contains("return"))
                return "Return";
            if (lower.Contains("phone") || lower.Contains("contact") || lower.Contains("unreachable"))
                return "Customer Not Picking Phone";

            return status; // Return as-is if already normalized
        }

        /// <summary>
        /// Normalize color to valid values
        /// </summary>
        private static string NormalizeColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color))
                return "Yellow";

            var lower = color.ToLower();

            if (lower.Contains("green")) return "Green";
            if (lower.Contains("yellow")) return "Yellow";
            if (lower.Contains("orange")) return "Orange";
            if (lower.Contains("red")) return "Red";

            return "Yellow"; // Default to yellow if unknown
        }

        /// <summary>
        /// Unload the model and free memory
        /// </summary>
        public void UnloadModel()
        {
            _executor = null;
            _context?.Dispose();
            _model?.Dispose();

            _context = null;
            _model = null;
            _isLoaded = false;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}