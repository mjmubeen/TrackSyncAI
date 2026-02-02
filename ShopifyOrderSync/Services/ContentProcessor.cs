using Newtonsoft.Json.Linq;
using ShopifyOrderSync.Enums;
using System.Text;
using System.Text.RegularExpressions;

namespace ShopifyOrderSync.Services
{
    /// <summary>
    /// Handles content processing and extraction for different content types
    /// </summary>
    public static class ContentProcessor
    {
        /// <summary>
        /// Detect what type of content we're dealing with
        /// </summary>
        public static ContentType DetectContentType(string content)
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
        public static string PrepareContentByType(string content, ContentType type)
        {
            // Emergency limit - never send more than 2000 chars to AI
            if (content.Length > 2000)
            {
                System.Diagnostics.Debug.WriteLine($"[CONTENT PROCESSOR] Input too large ({content.Length} chars), pre-truncating");
            }

            var result = type switch
            {
                ContentType.JSON => ProcessJSON(content),
                ContentType.XML => ProcessXML(content),
                ContentType.HTML => ProcessHTML(content),
                ContentType.PlainText => ProcessPlainText(content),
                _ => content
            };

            // Final safety check
            if (result.Length > 2000)
            {
                System.Diagnostics.Debug.WriteLine($"[CONTENT PROCESSOR] Output still too large ({result.Length} chars), force truncating");
                result = result[..2000];
            }

            System.Diagnostics.Debug.WriteLine($"[CONTENT PROCESSOR] Final: {content.Length} → {result.Length} chars");
            return result;
        }

        /// <summary>
        /// Process JSON API response - Extract only tracking-relevant fields
        /// </summary>
        public static string ProcessJSON(string json)
        {
            System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Starting - Input: {json.Length} chars");

            try
            {
                var jToken = JToken.Parse(json);
                var extracted = new StringBuilder();

                System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Parsed as: {jToken.Type}");

                // Pre-calculate HashSets ONCE
                var criticalSet = new HashSet<string>([
                    "status", "delivery_status", "tracking_status", "shipment_status", "current_status",
                    "order_status", "state", "stage", "step", "delivered", "is_delivered",
                    "deliveryStatus", "OperationDesc", "ProcessDescForPortal", "Status", "StatusCode",
                    "TrackingStatus", "CurrentStatus"
                ], StringComparer.OrdinalIgnoreCase);

                var locationSet = new HashSet<string>([
                    "location", "current_location", "last_location", "city", "ConsigneeCity",
                    "destination", "origin", "hub", "facility", "OriginCity", "DestBranch", "BranchName",
                    "CurrentLocation", "DestinationCity"
                ], StringComparer.OrdinalIgnoreCase);

                var timeSet = new HashSet<string>([
                    "date", "timestamp", "updated_at", "delivery_date", "TransactionDate",
                    "expected_delivery", "estimated_delivery", "delivered_at", "CallDate", "CallTime",
                    "DeliveryDate", "DateTime", "Time"
                ], StringComparer.OrdinalIgnoreCase);

                var contextSet = new HashSet<string>([
                    "remarks", "message", "description", "details", "notes",
                    "reason", "comment", "failed_reason", "exception", "ReasonDesc", "ConsigneeName",
                    "ReceivedBy", "Recipient"
                ], StringComparer.OrdinalIgnoreCase);

                // Check if root is an array (tracking history)
                if (jToken is JArray jArray && jArray.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Array with {jArray.Count} items");

                    // --- LATEST EVENT (primary focus) ---
                    var latest = jArray.Last;
                    extracted.AppendLine("### LATEST STATUS ###");
                    ExtractFields(latest, criticalSet, extracted, "STATUS");
                    ExtractFields(latest, locationSet, extracted, "LOCATION");
                    ExtractFields(latest, timeSet, extracted, "TIME");
                    ExtractFields(latest, contextSet, extracted, "DETAILS");

                    // --- HISTORY (last 3 events only) ---
                    if (jArray.Count > 1)
                    {
                        extracted.AppendLine("\n### RECENT HISTORY ###");
                        var history = jArray.Reverse().Skip(1).Take(3);
                        foreach (var item in history)
                        {
                            ExtractFields(item, criticalSet, extracted, "");
                            ExtractFields(item, timeSet, extracted, "");
                            extracted.AppendLine("---");
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Object structure");

                    // Regular object structure
                    ExtractFields(jToken, criticalSet, extracted, "STATUS");
                    ExtractFields(jToken, locationSet, extracted, "LOCATION");
                    ExtractFields(jToken, timeSet, extracted, "TIME");
                    ExtractFields(jToken, contextSet, extracted, "DETAILS");

                    // Try to find nested arrays (tracking history)
                    ExtractTrackingHistory(jToken, extracted);
                }

                var result = extracted.ToString();
                System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Extracted: {result.Length} chars");

                // Log first 300 chars of extracted content
                if (result.Length > 0)
                {
                    var preview = result.Length > 300 ? result[..300] + "..." : result;
                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Preview:\n{preview}");
                }

                // If nothing extracted, try pattern-based fallback
                if (string.IsNullOrWhiteSpace(result) || result.Length < 50)
                {
                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Extraction too small, trying pattern fallback");

                    var fallback = ExtractWithPatterns(json);
                    if (!string.IsNullOrWhiteSpace(fallback) && fallback.Length > 50)
                    {
                        System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Pattern fallback successful: {fallback.Length} chars");
                        return fallback.Length > 1500 ? fallback[..1500] : fallback;
                    }

                    // Ultimate fallback: Return cleaned JSON (limited)
                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Using cleaned JSON fallback");
                    return CleanJSON(json);
                }

                // Limit to 1500 chars
                var final = result.Length > 1500 ? result[..1500] : result;
                System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Success - Final: {final.Length} chars");
                return final;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Error: {ex.Message}");

                // Try pattern extraction
                var fallback = ExtractWithPatterns(json);
                return !string.IsNullOrWhiteSpace(fallback) ? fallback : ProcessPlainText(json);
            }
        }

        /// <summary>
        /// Extract fields from JSON token using HashSet
        /// </summary>
        private static void ExtractFields(JToken token, HashSet<string> fieldSet, StringBuilder output, string category)
        {
            var seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void WalkToken(JToken current)
            {
                if (current is JObject obj)
                {
                    foreach (var property in obj.Properties())
                    {
                        if (fieldSet.Contains(property.Name))
                        {
                            var value = property.Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(value) && seenValues.Add(value))
                            {
                                if (string.IsNullOrEmpty(category))
                                {
                                    output.AppendLine($"{property.Name}: {value}");
                                }
                                else
                                {
                                    output.AppendLine($"[{category}] {property.Name}: {value}");
                                }
                            }
                        }

                        if (property.Value.HasValues)
                        {
                            WalkToken(property.Value);
                        }
                    }
                }
                else if (current is JArray arr)
                {
                    foreach (var item in arr)
                    {
                        WalkToken(item);
                    }
                }
            }

            WalkToken(token);
        }

        /// <summary>
        /// Extract tracking history from nested arrays
        /// </summary>
        private static void ExtractTrackingHistory(JToken token, StringBuilder output)
        {
            var historyFields = new[] { "history", "events", "tracking_history", "shipment_history", "timeline", "updates", "History", "Events" };

            foreach (var field in historyFields)
            {
                var events = token.SelectTokens($"..{field}").FirstOrDefault();

                if (events?.Type == JTokenType.Array)
                {
                    var eventArray = (JArray)events;

                    System.Diagnostics.Debug.WriteLine($"[JSON PROCESSOR] Found history array '{field}' with {eventArray.Count} events");

                    // Get last 3 events (most recent)
                    var recentEvents = eventArray.Reverse().Take(3).Reverse();

                    output.AppendLine("\n### TRACKING HISTORY ###");
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
        /// Fallback: Extract using regex patterns when field extraction fails
        /// </summary>
        private static string ExtractWithPatterns(string json)
        {
            System.Diagnostics.Debug.WriteLine($"[PATTERN EXTRACTOR] Starting pattern-based extraction");

            var sb = new StringBuilder();

            // Pattern 1: "status": "value"
            var statusPattern = @"""(?:status|delivery_status|tracking_status|state|stage|ProcessDescForPortal|OperationDesc)""\s*:\s*""([^""]+)""";
            var statusMatches = Regex.Matches(json, statusPattern, RegexOptions.IgnoreCase);

            System.Diagnostics.Debug.WriteLine($"[PATTERN EXTRACTOR] Found {statusMatches.Count} status matches");

            foreach (Match match in statusMatches.Take(5))
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    sb.AppendLine($"[STATUS] {match.Groups[1].Value}");
                }
            }

            // Pattern 2: "location": "value" or "city": "value"
            var locationPattern = @"""(?:location|city|destination|origin|hub|BranchName|ConsigneeCity)""\s*:\s*""([^""]+)""";
            var locationMatches = Regex.Matches(json, locationPattern, RegexOptions.IgnoreCase);

            System.Diagnostics.Debug.WriteLine($"[PATTERN EXTRACTOR] Found {locationMatches.Count} location matches");

            foreach (Match match in locationMatches.Take(3))
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    sb.AppendLine($"[LOCATION] {match.Groups[1].Value}");
                }
            }

            // Pattern 3: "date": "value" or "timestamp": "value"
            var datePattern = @"""(?:date|timestamp|delivered_at|delivery_date|TransactionDate)""\s*:\s*""([^""]+)""";
            var dateMatches = Regex.Matches(json, datePattern, RegexOptions.IgnoreCase);

            System.Diagnostics.Debug.WriteLine($"[PATTERN EXTRACTOR] Found {dateMatches.Count} date matches");

            foreach (Match match in dateMatches.Take(2))
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    sb.AppendLine($"[TIME] {match.Groups[1].Value}");
                }
            }

            var result = sb.ToString();
            System.Diagnostics.Debug.WriteLine($"[PATTERN EXTRACTOR] Extracted: {result.Length} chars");

            return result;
        }

        /// <summary>
        /// Clean JSON by removing unnecessary whitespace
        /// </summary>
        private static string CleanJSON(string json)
        {
            var cleaned = Regex.Replace(json, @"\s+", " ");
            return cleaned.Length > 1000 ? cleaned[..1000] : cleaned;
        }

        /// <summary>
        /// Process XML response
        /// </summary>
        public static string ProcessXML(string xml)
        {
            try
            {
                var pattern = @"<(?:status|location|date|time|message|description|remarks|delivery)[^>]*>([^<]+)</";
                var matches = Regex.Matches(xml, pattern, RegexOptions.IgnoreCase);

                var sb = new StringBuilder();
                foreach (Match match in matches)
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
        /// Process HTML page
        /// </summary>
        public static string ProcessHTML(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return "";

            // Remove script and style tags
            html = Regex.Replace(html, @"<script[^>]*>.*?</script>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"<style[^>]*>.*?</style>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"<!--.*?-->", "");

            // Remove all HTML tags
            html = Regex.Replace(html, @"<[^>]+>", " ");

            // Decode HTML entities
            html = System.Net.WebUtility.HtmlDecode(html);

            // Remove extra whitespace
            html = Regex.Replace(html, @"\s+", " ").Trim();

            return TruncateIntelligently(html, 1500);
        }

        /// <summary>
        /// Process plain text
        /// </summary>
        public static string ProcessPlainText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = Regex.Replace(text, @"\s+", " ").Trim();
            return TruncateIntelligently(text, 1500);
        }

        /// <summary>
        /// Intelligent truncation based on keywords
        /// </summary>
        private static string TruncateIntelligently(string text, int maxLength)
        {
            if (text.Length <= maxLength)
                return text;

            var keywords = new[]
            {
                "delivered", "delivery", "status", "tracking",
                "failed", "returned", "return", "stuck", "transit",
                "location", "date", "time", "received", "recipient",
                "out for delivery", "in transit", "picked up",
                "attempted", "exception", "delay", "completed"
            };

            var segments = text.Split(['.', '\n', ';'], StringSplitOptions.RemoveEmptyEntries);
            var importantSegments = new List<string>();
            var currentLength = 0;

            foreach (var segment in segments)
            {
                var lowerSegment = segment.ToLower();
                if (keywords.Any(k => lowerSegment.Contains(k)))
                {
                    if (currentLength + segment.Length < maxLength)
                    {
                        importantSegments.Add(segment.Trim());
                        currentLength += segment.Length + 2;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            if (importantSegments.Count > 0 && currentLength > maxLength / 2)
            {
                return string.Join(". ", importantSegments);
            }

            var halfLength = maxLength / 2 - 10;
            var beginning = text[..Math.Min(halfLength, text.Length)];
            var endStart = Math.Max(halfLength, text.Length - halfLength);
            var end = text[endStart..];

            return beginning + " [...] " + end;
        }
    }
}