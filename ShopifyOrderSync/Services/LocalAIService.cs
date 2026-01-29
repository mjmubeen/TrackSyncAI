using LLama;
using LLama.Common;
using LLama.Sampling;
using Newtonsoft.Json;
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
        /// Analyze tracking HTML content using the AI model
        /// </summary>
        public async Task<TrackingAnalysisResult> AnalyzeTrackingAsync(string trackingHtml)
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
                    // Clean HTML (remove scripts, styles, trim whitespace)
                    string cleanedText = CleanHtmlForAnalysis(trackingHtml);

                    // Limit text size to prevent token overflow
                    if (cleanedText.Length > 3000)
                    {
                        cleanedText = cleanedText[..3000];
                    }

                    // Construct prompt for tracking analysis
                    string prompt = $@"<|system|>
You are a courier tracking analyzer for an e-commerce business in Pakistan. Analyze tracking info and detect problems.
Return ONLY valid JSON: {{""status"": ""Status"", ""color"": ""Color""}}

Status options:
- Delivered: Package successfully delivered
- In-Transit: Moving normally through courier network
- Stuck: No movement for 2+ days at same location
- Failed: Delivery attempt failed
- Return: Being returned to sender
- Customer Not Picking Phone: Courier cannot contact customer

Color codes:
- Green: Delivered
- Yellow: In-Transit (normal)
- Orange: Stuck (warning - needs follow-up)
- Red: Failed, Return, or Customer not reachable (urgent action needed)

Look for keywords like: delivered, out for delivery, in transit, attempted delivery, returned, customer unreachable, contact failed, stuck, delay
<|end|>
<|user|>
Analyze this courier tracking information:

{cleanedText}

Return JSON with status and color.<|end|>
<|assistant|>";

                    var inferenceParams = new InferenceParams
                    {
                        SamplingPipeline = new DefaultSamplingPipeline() { Temperature = 0.3f, Seed = 1337 },
                        AntiPrompts = ["<|end|>", "<|user|>"],
                        MaxTokens = 150
                    };

                    // Use InferAsync for async operation
                    string response = "";
                    await foreach (var token in _executor.InferAsync(prompt, inferenceParams))
                    {
                        response += token;
                    }

                    // Extract JSON from response
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
        /// Clean HTML text for AI analysis
        /// </summary>
        private static string CleanHtmlForAnalysis(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return "";

            // Remove script and style tags
            html = System.Text.RegularExpressions.Regex.Replace(
                html, @"<script[^>]*>.*?</script>", "",
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            html = System.Text.RegularExpressions.Regex.Replace(
                html, @"<style[^>]*>.*?</style>", "",
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Remove all HTML tags
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<[^>]+>", " ");

            // Decode HTML entities
            html = System.Net.WebUtility.HtmlDecode(html);

            // Remove extra whitespace
            html = System.Text.RegularExpressions.Regex.Replace(html, @"\s+", " ").Trim();

            return html;
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
                    var result = JsonConvert.DeserializeObject<TrackingAnalysisResult>(json) ?? throw new Exception("Data Not Found");

                    // Validate result
                    if (string.IsNullOrWhiteSpace(result.Status))
                        result.Status = "In-Transit";
                    if (string.IsNullOrWhiteSpace(result.Color))
                        result.Color = "Yellow";

                    return result;
                }

                // Fallback: Try to parse entire response as JSON
                return JsonConvert.DeserializeObject<TrackingAnalysisResult>(aiResponse) ?? throw new Exception("Data Not Found");
            }
            catch
            {
                // Fallback result
                return new TrackingAnalysisResult
                {
                    Status = "In-Transit",
                    Color = "Yellow",
                    ErrorMessage = "Could not parse AI response"
                };
            }
        }

        /// <summary>
        /// Unload the model and free memory
        /// </summary>
        public void UnloadModel()
        {
            // Note: InteractiveExecutor doesn't implement IDisposable in current version
            // We only dispose the context and model
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