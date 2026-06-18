using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    internal sealed class PowerPointPdfPreviewControl : UserControl, IDisposable
    {
        private static readonly string CacheRoot = Path.Combine(Path.GetTempPath(), "QuickLook.PowerPointNativeViewer");
        private static readonly string PersistentCacheRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "QuickLook.PowerPointNativeViewer",
            "pdf-cache");

        private readonly string _path;
        private readonly string _sessionDir;
        private readonly string _cachedPdfPath;
        private readonly TextBlock _status;
        private readonly WebView2 _webView;
        private bool _disposed;

        public PowerPointPdfPreviewControl(string path)
        {
            _path = path;
            _sessionDir = Path.Combine(CacheRoot, "session-" + Guid.NewGuid().ToString("N"));
            _cachedPdfPath = GetCachedPdfPath(path);
            Directory.CreateDirectory(_sessionDir);
            Directory.CreateDirectory(Path.GetDirectoryName(_cachedPdfPath));

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
            Loaded += delegate
            {
                BringHostWindowToFront();
                Start();
            };
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

            try
            {
                if (!Directory.Exists(PersistentCacheRoot))
                    return;

                foreach (var dir in Directory.GetDirectories(PersistentCacheRoot))
                {
                    try
                    {
                        if (Directory.GetLastWriteTimeUtc(dir) < DateTime.UtcNow.AddDays(-14))
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

                if (File.Exists(_cachedPdfPath))
                {
                    await ShowPdfAsync(_cachedPdfPath, "Opening cached PDF preview...");
                    return;
                }

                SetStatus("Generating PDF preview...");
                var pdfPath = await Task.Run(delegate
                {
                    using (var session = new PowerPointRenderSession(_path, _sessionDir))
                    {
                        var sessionPdf = session.ExportPdf();
                        File.Copy(sessionPdf, _cachedPdfPath, true);
                        return _cachedPdfPath;
                    }
                });

                if (_disposed)
                    return;

                await ShowPdfAsync(pdfPath, "Opening PDF preview...");
            }
            catch (Exception ex)
            {
                if (!_disposed)
                    SetStatus("PowerPoint PDF preview failed: " + ex.Message);
            }
        }

        private async Task ShowPdfAsync(string pdfPath, string status)
        {
            SetStatus(status);
            await InitializeWebViewAsync();

            if (_disposed)
                return;

            _webView.Source = new Uri(pdfPath);
            _webView.Visibility = Visibility.Visible;
            _status.Visibility = Visibility.Collapsed;
            BringHostWindowToFront();
        }

        private async Task InitializeWebViewAsync()
        {
            if (_webView.CoreWebView2 != null)
                return;

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

        private void BringHostWindowToFront()
        {
            Dispatcher.BeginInvoke(new Action(delegate
            {
                try
                {
                    var window = Window.GetWindow(this);
                    if (window == null)
                        return;

                    if (window.WindowState == WindowState.Minimized)
                        window.WindowState = WindowState.Normal;

                    window.Topmost = true;
                    window.Activate();
                    window.Focus();
                    window.Topmost = false;
                }
                catch
                {
                }
            }), DispatcherPriority.ApplicationIdle);
        }

        private static string GetCachedPdfPath(string path)
        {
            var info = new FileInfo(path);
            var key = string.Format(
                CultureInfo.InvariantCulture,
                "{0}|{1}|{2}",
                info.FullName.ToLowerInvariant(),
                info.Length,
                info.LastWriteTimeUtc.Ticks);

            var hash = Sha256(key);
            return Path.Combine(PersistentCacheRoot, hash, "preview.pdf");
        }

        private static string Sha256(string value)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (var b in bytes)
                    builder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }
    }
}
