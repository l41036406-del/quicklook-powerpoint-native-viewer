using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    internal sealed class PowerPointPdfPreviewControl : UserControl, IDisposable
    {
        private static readonly string CacheRoot = Path.Combine(Path.GetTempPath(), "QuickLook.PowerPointNativeViewer");

        private readonly string _path;
        private readonly string _sessionDir;
        private readonly TextBlock _status;
        private readonly WebView2 _webView;
        private bool _disposed;

        public PowerPointPdfPreviewControl(string path)
        {
            _path = path;
            _sessionDir = Path.Combine(CacheRoot, "session-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_sessionDir);

            var root = new Grid { Background = Brushes.Black };

            _webView = new WebView2
            {
                Visibility = Visibility.Collapsed
            };

            _status = new TextBlock
            {
                Text = "Starting PowerPoint...",
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromArgb(170, 0, 0, 0)),
                Padding = new Thickness(12),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            root.Children.Add(_webView);
            root.Children.Add(_status);

            Content = root;
            Loaded += delegate { Start(); };
            Unloaded += delegate { Dispose(); };
        }

        public static void CleanupOldSessions()
        {
            try
            {
                if (!Directory.Exists(CacheRoot))
                    return;

                foreach (var dir in Directory.GetDirectories(CacheRoot, "session-*"))
                {
                    try
                    {
                        if (Directory.GetLastWriteTimeUtc(dir) < DateTime.UtcNow.AddDays(-1))
                            Directory.Delete(dir, true);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            try
            {
                _webView.Dispose();
            }
            catch
            {
            }

            try
            {
                if (Directory.Exists(_sessionDir))
                    Directory.Delete(_sessionDir, true);
            }
            catch
            {
            }
        }

        private async void Start()
        {
            if (_disposed)
                return;

            try
            {
                SetStatus("Starting PowerPoint...");
                var pdfPath = await Task.Run(delegate
                {
                    using (var session = new PowerPointRenderSession(_path, _sessionDir))
                    {
                        return session.ExportPdf();
                    }
                });

                if (_disposed)
                    return;

                SetStatus("Opening PDF preview...");
                await InitializeWebViewAsync();

                if (_disposed)
                    return;

                _webView.Source = new Uri(pdfPath);
                _webView.Visibility = Visibility.Visible;
                _status.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                if (!_disposed)
                    SetStatus("PowerPoint PDF preview failed: " + ex.Message);
            }
        }

        private async Task InitializeWebViewAsync()
        {
            var userDataFolder = Path.Combine(_sessionDir, "WebView2");
            Directory.CreateDirectory(userDataFolder);
            var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await _webView.EnsureCoreWebView2Async(environment);
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
        }

        private void SetStatus(string text)
        {
            _status.Text = text;
            _status.Visibility = Visibility.Visible;
        }
    }
}
