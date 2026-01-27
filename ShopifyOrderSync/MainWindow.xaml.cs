using Microsoft.Win32;
using Newtonsoft.Json;
using ShopifyOrderSync.Services;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;

namespace ShopifyOrderSync
{
    public partial class MainWindow : Window
    {
        private LocalAIService _aiService;
        private OrderSyncService _syncService;
        private DispatcherTimer _memoryTimer;

        public MainWindow()
        {
            InitializeComponent();
            InitializeApp();
        }

        private void InitializeApp()
        {
            _aiService = LocalAIService.Instance;

            // Set default dates
            StartDatePicker.SelectedDate = DateTime.Now.AddDays(-30);
            EndDatePicker.SelectedDate = DateTime.Now;

            // Setup memory monitor
            _memoryTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            _memoryTimer.Tick += UpdateMemoryUsage;
            _memoryTimer.Start();

            Log("Application initialized. Please load the AI model before syncing.");
        }

        private void BrowseModel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "GGUF Model Files (*.gguf)|*.gguf|All Files (*.*)|*.*",
                Title = "Select AI Model File"
            };

            if (dialog.ShowDialog() == true)
            {
                ModelPathTextBox.Text = dialog.FileName;
            }
        }

        private async void LoadModel_Click(object sender, RoutedEventArgs e)
        {
            string modelPath = ModelPathTextBox.Text;

            if (string.IsNullOrWhiteSpace(modelPath) || !File.Exists(modelPath))
            {
                MessageBox.Show("Please select a valid model file.", "Invalid Path",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            LoadModelButton.IsEnabled = false;
            LoadModelButton.Content = "Loading...";
            ModelStatusTextBlock.Text = "Loading model...";
            ModelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Orange;

            Log("\n=== Loading AI Model ===");

            bool success = await _aiService.LoadModelAsync(modelPath, Log);

            if (success)
            {
                ModelStatusTextBlock.Text = "Model Loaded ✓";
                ModelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                LoadModelButton.Content = "Reload Model";
                LoadModelButton.IsEnabled = true;

                MessageBox.Show("AI Model loaded successfully!\n\nYou can now start syncing orders.",
                    "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                ModelStatusTextBlock.Text = "Load Failed ✗";
                ModelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                LoadModelButton.Content = "Retry Load";
                LoadModelButton.IsEnabled = true;

                MessageBox.Show("Failed to load the AI model. Please check the log for details.",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void SyncButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate AI model is loaded
            if (!_aiService.IsLoaded)
            {
                MessageBox.Show("Please load the AI model before syncing.", "Model Not Loaded",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Validate dates
            if (!StartDatePicker.SelectedDate.HasValue || !EndDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Please select both start and end dates.", "Invalid Dates",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Get configuration (in production, load from config file or settings)
            var config = LoadConfiguration();
            if (config == null)
            {
                MessageBox.Show("Configuration not found. Please set up your Shopify and Google credentials.",
                    "Configuration Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Initialize sync service
            _syncService = new OrderSyncService(
                config.ShopifyApiKey,
                config.ShopifyPassword,
                config.ShopifyShopDomain,
                System.Text.Json.JsonSerializer.Serialize(config.GoogleCredentialsJson),
                config.SpreadsheetId
            );

            _syncService.LogEvent += Log;
            _syncService.ProgressEvent += UpdateProgress;

            // Disable UI during sync
            LoadModelButton.IsEnabled = false;
            SyncButton.IsEnabled = false;
            SyncButton.Content = "Syncing...";
            StartDatePicker.IsEnabled = false;
            EndDatePicker.IsEnabled = false;

            try
            {
                await _syncService.SyncOrdersAsync(
                    StartDatePicker.SelectedDate.Value,
                    EndDatePicker.SelectedDate.Value
                );

                MessageBox.Show("Sync completed successfully!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (System.Net.Http.HttpRequestException httpEx)
            {
                string errorMsg = $"HTTP Connection Error:\n\n{httpEx.Message}\n\n";

                if (httpEx.Message.Contains("401") || httpEx.Message.Contains("Unauthorized"))
                {
                    errorMsg += "LIKELY CAUSE: Invalid Shopify API token or wrong shop domain.\n\n" +
                                "Check your config.json:\n" +
                                "- ShopifyPassword must start with 'shpat_'\n" +
                                "- ShopifyShopDomain must be 'yourstore.myshopify.com'";
                }
                else if (httpEx.Message.Contains("403") || httpEx.Message.Contains("Forbidden"))
                {
                    errorMsg += "LIKELY CAUSE: Google Sheets permission denied.\n\n" +
                                "Check:\n" +
                                "- Share sheet with service account email\n" +
                                "- Give Editor permissions\n" +
                                "- Verify Sheet ID in config.json";
                }
                else if (httpEx.Message.Contains("404") || httpEx.Message.Contains("Not Found"))
                {
                    errorMsg += "LIKELY CAUSE: Wrong URL or resource not found.\n\n" +
                                "Check:\n" +
                                "- Shopify shop domain is correct\n" +
                                "- Google Sheet ID is correct";
                }
                else
                {
                    errorMsg += "LIKELY CAUSE: Network connection issue.\n\n" +
                                "Check:\n" +
                                "- Internet connection working\n" +
                                "- Firewall/antivirus not blocking\n" +
                                "- VPN not interfering";
                }

                MessageBox.Show(errorMsg, "Connection Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Log($"HTTP ERROR: {httpEx.Message}");
                Log($"Stack Trace: {httpEx.StackTrace}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Sync failed: {ex.Message}\n\nType: {ex.GetType().Name}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Log($"SYNC FAILED: {ex.Message}");
                Log($"Exception Type: {ex.GetType().Name}");
                Log($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                // Re-enable UI
                LoadModelButton.IsEnabled = true;
                SyncButton.IsEnabled = true;
                SyncButton.Content = "Start Sync";
                StartDatePicker.IsEnabled = true;
                EndDatePicker.IsEnabled = true;
                ProgressBar.Value = 0;
                ProgressTextBlock.Text = "0%";
            }
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                LogTextBox.ScrollToEnd();
            });
        }

        private void UpdateProgress(int percent)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = percent;
                ProgressTextBlock.Text = $"{percent}%";
                StatusTextBlock.Text = $"Processing... {percent}% complete";
            });
        }

        private void UpdateMemoryUsage(object? sender, EventArgs e)
        {
            long memoryMB = _aiService.MemoryUsageMB;
            MemoryUsageTextBlock.Text = $"{memoryMB} MB";
        }

        private static AppConfiguration? LoadConfiguration()
        {
            // In production, load this from app.config, appsettings.json, or encrypted storage
            // For demo purposes, return a sample configuration

            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

            if (File.Exists(configPath))
            {
                try
                {
                    string json = File.ReadAllText(configPath);
                    return Newtonsoft.Json.JsonConvert.DeserializeObject<AppConfiguration>(json);
                }
                catch
                {
                    return null;
                }
            }

            // Show configuration dialog if not found
            MessageBox.Show(
                "Configuration file not found!\n\n" +
                "Please create a 'config.json' file in the application directory with:\n" +
                "{\n" +
                "  \"ShopifyApiKey\": \"your-api-key\",\n" +
                "  \"ShopifyPassword\": \"your-password\",\n" +
                "  \"ShopifyShopDomain\": \"yourstore.myshopify.com\",\n" +
                "  \"GoogleCredentialsJson\": \"{...}\",\n" +
                "  \"SpreadsheetId\": \"your-sheet-id\"\n" +
                "}",
                "Configuration Required",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );

            return null;
        }

        protected override void OnClosed(EventArgs e)
        {
            _memoryTimer?.Stop();
            _aiService?.UnloadModel();
            base.OnClosed(e);
        }
    }

    public class AppConfiguration
    {
        public required string ShopifyApiKey { get; set; }
        public required string ShopifyPassword { get; set; }
        public required string ShopifyShopDomain { get; set; }
        public required GoogleCredentials GoogleCredentialsJson { get; set; }
        public required string SpreadsheetId { get; set; }
    }
    public class GoogleCredentials
    {
        [JsonProperty("type")]
        public required string Type { get; set; }

        [JsonProperty("project_id")]
        public required string ProjectId { get; set; }

        [JsonProperty("private_key_id")]
        public required string PrivateKeyId { get; set; }

        [JsonProperty("private_key")]
        public required string PrivateKey { get; set; }

        [JsonProperty("client_email")]
        public required string ClientEmail { get; set; }

        [JsonProperty("client_id")]
        public required string ClientId { get; set; }

        [JsonProperty("auth_uri")]
        public required string AuthUri { get; set; }

        [JsonProperty("token_uri")]
        public required string TokenUri { get; set; }

        [JsonProperty("auth_provider_x509_cert_url")]
        public required string AuthProviderX509CertUrl { get; set; }

        [JsonProperty("client_x509_cert_url")]
        public required string ClientX509CertUrl { get; set; }

        [JsonProperty("universe_domain")]
        public required string UniverseDomain { get; set; }

    }
}