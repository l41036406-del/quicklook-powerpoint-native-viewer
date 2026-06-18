using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    internal sealed class PowerPointPreviewControl : UserControl, IDisposable
    {
        private static readonly string CacheRoot = Path.Combine(Path.GetTempPath(), "QuickLook.PowerPointNativeViewer");

        private readonly string _path;
        private readonly string _sessionDir;
        private readonly Image _image;
        private readonly TextBlock _status;
        private readonly TextBlock _pageLabel;
        private readonly Button _previousButton;
        private readonly Button _nextButton;
        private PowerPointRenderSession _session;
        private int _currentPage = 1;
        private int _requestId;
        private bool _disposed;

        public PowerPointPreviewControl(string path)
        {
            _path = path;
            _sessionDir = Path.Combine(CacheRoot, "session-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_sessionDir);

            var root = new DockPanel();
            Background = Brushes.Black;

            var toolbar = new Grid { Height = 42, Background = new SolidColorBrush(Color.FromArgb(230, 32, 32, 32)) };
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            _previousButton = CreateButton("‹");
            _nextButton = CreateButton("›");
            _pageLabel = new TextBlock
            {
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(14, 0, 0, 0)
            };

            _previousButton.Click += delegate { MovePage(-1); };
            _nextButton.Click += delegate { MovePage(1); };

            Grid.SetColumn(_previousButton, 0);
            Grid.SetColumn(_nextButton, 1);
            Grid.SetColumn(_pageLabel, 2);
            toolbar.Children.Add(_previousButton);
            toolbar.Children.Add(_nextButton);
            toolbar.Children.Add(_pageLabel);

            DockPanel.SetDock(toolbar, Dock.Bottom);
            root.Children.Add(toolbar);

            var imageHost = new Grid();
            _image = new Image { Stretch = Stretch.Uniform };
            _status = new TextBlock
            {
                Text = "Loading PowerPoint preview...",
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromArgb(170, 0, 0, 0)),
                Padding = new Thickness(12),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            imageHost.Children.Add(_image);
            imageHost.Children.Add(_status);
            root.Children.Add(imageHost);

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

            if (_session != null)
            {
                _session.Dispose();
                _session = null;
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

        private static Button CreateButton(string text)
        {
            return new Button
            {
                Content = text,
                Width = 44,
                Height = 32,
                Margin = new Thickness(6, 5, 0, 5),
                FontSize = 22,
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromRgb(55, 55, 55)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(95, 95, 95))
            };
        }

        private async void Start()
        {
            if (_session != null || _disposed)
                return;

            SetStatus("Starting PowerPoint...");
            try
            {
                _session = await Task.Run(delegate { return new PowerPointRenderSession(_path, _sessionDir); });
                UpdateButtons();
                await RenderPageAsync(1);
            }
            catch (Exception ex)
            {
                SetStatus("PowerPoint preview failed: " + ex.Message);
                UpdateButtons();
            }
        }

        private async void MovePage(int delta)
        {
            if (_session == null || _disposed)
                return;

            var target = _currentPage + delta;
            if (target < 1 || target > _session.SlideCount)
                return;

            await RenderPageAsync(target);
        }

        private async Task RenderPageAsync(int page)
        {
            if (_session == null || _disposed)
                return;

            var request = ++_requestId;
            SetStatus("Rendering page " + page + "...");
            UpdateButtons();

            try
            {
                var size = GetExportSize();
                var path = await Task.Run(delegate { return _session.ExportPage(page, size.Width, size.Height); });

                if (_disposed || request != _requestId)
                    return;

                _image.Source = LoadBitmap(path);
                _currentPage = page;
                _status.Visibility = Visibility.Collapsed;
                UpdateButtons();
                PreRenderNeighbor(page + 1);
            }
            catch (Exception ex)
            {
                if (!_disposed && request == _requestId)
                    SetStatus("Render failed: " + ex.Message);
            }
        }

        private void PreRenderNeighbor(int page)
        {
            if (_session == null || _disposed || page < 1 || page > _session.SlideCount)
                return;

            var size = GetExportSize();
            Task.Run(delegate
            {
                try
                {
                    _session.ExportPage(page, size.Width, size.Height);
                }
                catch
                {
                }
            });
        }

        private ExportSize GetExportSize()
        {
            var targetWidth = Math.Max(900, Math.Min(1920, (int)Math.Round(Math.Max(ActualWidth, 1280) * 1.4)));
            var aspect = 0.5625;
            if (_session != null && _session.SlideWidth > 0)
                aspect = _session.SlideHeight / _session.SlideWidth;

            var targetHeight = Math.Max(1, (int)Math.Round(targetWidth * aspect));
            return new ExportSize(targetWidth, targetHeight);
        }

        private static BitmapImage LoadBitmap(string path)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(path);
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        private void SetStatus(string text)
        {
            _status.Text = text;
            _status.Visibility = Visibility.Visible;
        }

        private void UpdateButtons()
        {
            var slideCount = _session == null ? 0 : _session.SlideCount;
            _previousButton.IsEnabled = _session != null && _currentPage > 1;
            _nextButton.IsEnabled = _session != null && _currentPage < slideCount;
            _pageLabel.Text = slideCount == 0 ? string.Empty : string.Format("{0} / {1}", _currentPage, slideCount);
        }

        private struct ExportSize
        {
            public readonly int Width;
            public readonly int Height;

            public ExportSize(int width, int height)
            {
                Width = width;
                Height = height;
            }
        }
    }
}
