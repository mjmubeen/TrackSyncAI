using LLama;
using LLama.Common;
using LLama.Sampling;
using Newtonsoft.Json;
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
        private ModelParams? _parameters;
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
                    _parameters = new ModelParams(modelPath)
                    {
                        ContextSize = 4096,      // Context window
                        GpuLayerCount = 0,       // Use CPU only (set to > 0 for GPU)
                        Embeddings = false
                    };

                    logCallback?.Invoke("Initializing LLamaWeights...");
                    _model = LLamaWeights.LoadFromFile(_parameters);
                    
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
            if (!_isLoaded || _model == null || _parameters == null)
            {
                throw new InvalidOperationException("Model not loaded. Please load the model first.");
            }

            return await Task.Run(async () =>
            {
                LLamaContext? context = null;
                InteractiveExecutor? executor = null;
                try
                {
                    context = _model.CreateContext(_parameters);
                    executor = new InteractiveExecutor(context);

                    System.Diagnostics.Debug.WriteLine($"=== AI ANALYSIS START ===");
                    System.Diagnostics.Debug.WriteLine($"Input: {trackingContent.Length} chars");

                    // Use ContentProcessor to prepare content
                    var contentType = ContentProcessor.DetectContentType(trackingContent);
                    System.Diagnostics.Debug.WriteLine($"Type: {contentType}");

                    string preparedText = ContentProcessor.PrepareContentByType(trackingContent, contentType);
                    System.Diagnostics.Debug.WriteLine($"Prepared: {preparedText.Length} chars");

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
                        AntiPrompts = ["<|eot_id|>", "<|end_header_id|>"],
                        MaxTokens = 100
                    };

                    // Get AI response with early stopping
                    string response = "";
                    int tokenCount = 0;
                    await foreach (var token in executor.InferAsync(prompt, inferenceParams))
                    {
                        if (!string.IsNullOrEmpty(token))
                        {
                            response += token;
                            tokenCount++;
                        }

                        // Early stop if we have valid JSON
                        if (response.Contains('}') && response.Contains("status"))
                        {
                            break;
                        }

                        // Safety: Stop if response is too long
                        if (tokenCount > 150) break;
                    }

                    System.Diagnostics.Debug.WriteLine($"Response: {response}");

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
                    System.Diagnostics.Debug.WriteLine($"ERROR: {ex.Message}");
                    return new TrackingAnalysisResult
                    {
                        Status = "Error",
                        Color = "Red",
                        ErrorMessage = ex.Message
                    };
                }
                finally
                {
                    // CRITICAL: Dispose context and executor after EACH inference
                    executor = null;
                    context?.Dispose();
                    System.Diagnostics.Debug.WriteLine($"=== CONTEXT DISPOSED ===");
                }
            });
        }

        /// <summary>
        /// Detect what type of content we're dealing with
        /// </summary>
        private static string BuildPrompt(string preparedText, ContentType type)
        {
            var contentDescription = type switch
            {
                ContentType.JSON => "JSON API response",
                ContentType.XML => "XML tracking data",
                ContentType.HTML => "web page content",
                _ => "tracking information"
            };

            // Llama-3 format
            return $@"<|begin_of_text|><|start_header_id|>system<|end_header_id|>

You are a courier tracking analyzer. Analyze {contentDescription} and determine delivery status.
Return ONLY valid JSON: {{""status"": ""Status"", ""color"": ""Color""}}

Status options: Delivered, In-Transit, Stuck, Failed, Return, Customer Not Picking Phone
Color options: Green, Yellow, Orange, Red<|eot_id|><|start_header_id|>user<|end_header_id|>

{preparedText}

Return JSON:<|eot_id|><|start_header_id|>assistant<|end_header_id|>

";
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
            if (string.IsNullOrWhiteSpace(status)) return "In-Transit";
            var lower = status.ToLower();
            if (lower.Contains("deliver") && !lower.Contains("not")) return "Delivered";
            if (lower.Contains("transit")) return "In-Transit";
            if (lower.Contains("stuck") || lower.Contains("delay")) return "Stuck";
            if (lower.Contains("fail")) return "Failed";
            if (lower.Contains("return")) return "Return";
            if (lower.Contains("phone")) return "Customer Not Picking Phone";
            return status; // Return as-is if already normalized
        }

        /// <summary>
        /// Normalize color to valid values
        /// </summary>
        private static string NormalizeColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color)) return "Yellow";
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
            _model?.Dispose();

            _parameters = null;
            _model = null;
            _isLoaded = false;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}