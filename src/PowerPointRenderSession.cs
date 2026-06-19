using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    internal sealed class PowerPointRenderSession : IDisposable
    {
        private const int MsoTrue = -1;
        private const int MsoFalse = 0;

        private readonly BlockingCollection<Action> _queue = new BlockingCollection<Action>();
        private readonly ManualResetEventSlim _ready = new ManualResetEventSlim(false);
        private readonly string _path;
        private readonly string _sessionDir;
        private readonly Thread _thread;
        private dynamic _app;
        private dynamic _presentation;
        private Exception _startupException;
        private int _powerPointProcessId;
        private bool _disposed;

        public PowerPointRenderSession(string path, string sessionDir)
        {
            _path = path;
            _sessionDir = sessionDir;
            _thread = new Thread(Run);
            _thread.IsBackground = true;
            _thread.Name = "QuickLook PowerPoint renderer";
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            _ready.Wait();

            if (_startupException != null)
                throw _startupException;
        }

        public int SlideCount { get; private set; }

        public double SlideWidth { get; private set; }

        public double SlideHeight { get; private set; }

        public string ExportPage(int page, int width, int height)
        {
            if (page < 1 || page > SlideCount)
                throw new ArgumentOutOfRangeException("page");

            var fileName = string.Format("page-{0}-{1}x{2}.png", page, width, height);
            var outputPath = Path.Combine(_sessionDir, fileName);

            if (File.Exists(outputPath))
                return outputPath;

            return Invoke(delegate
            {
                dynamic slide = _presentation.Slides[page];
                slide.Export(outputPath, "PNG", width, height);
                ReleaseComObject(slide);
                return outputPath;
            });
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            _queue.CompleteAdding();
            if (!_thread.Join(TimeSpan.FromSeconds(5)))
                _thread.Join(TimeSpan.FromSeconds(1));
            _queue.Dispose();
            _ready.Dispose();
        }

        private T Invoke<T>(Func<T> func)
        {
            if (_disposed)
                throw new ObjectDisposedException("PowerPointRenderSession");

            var completion = new TaskCompletionSource<T>();
            _queue.Add(delegate
            {
                try
                {
                    completion.SetResult(func());
                }
                catch (Exception ex)
                {
                    completion.SetException(ex);
                }
            });

            return completion.Task.GetAwaiter().GetResult();
        }

        private void Run()
        {
            try
            {
                var type = Type.GetTypeFromProgID("PowerPoint.Application");
                if (type == null)
                    throw new InvalidOperationException("Microsoft PowerPoint is not installed or COM registration is missing.");

                _app = Activator.CreateInstance(type);
                CapturePowerPointProcessId();
                try
                {
                    _app.DisplayAlerts = MsoFalse;
                }
                catch
                {
                }

                _presentation = _app.Presentations.Open(_path, MsoTrue, MsoFalse, MsoFalse);
                SlideCount = (int)_presentation.Slides.Count;
                SlideWidth = Convert.ToDouble(_presentation.PageSetup.SlideWidth);
                SlideHeight = Convert.ToDouble(_presentation.PageSetup.SlideHeight);
            }
            catch (Exception ex)
            {
                _startupException = ex;
            }
            finally
            {
                _ready.Set();
            }

            if (_startupException == null)
            {
                foreach (var action in _queue.GetConsumingEnumerable())
                    action();
            }

            ClosePowerPoint();
        }

        private void ClosePowerPoint()
        {
            var presentation = (object)_presentation;
            _presentation = null;

            try
            {
                InvokeComMethod(presentation, "Close");
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(presentation);
            }

            var app = (object)_app;
            _app = null;

            try
            {
                InvokeComMethod(app, "Quit");
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(app);
                EnsurePowerPointProcessClosed();
            }
        }

        private void CapturePowerPointProcessId()
        {
            try
            {
                var hwnd = new IntPtr((int)_app.HWND);
                int processId;
                GetWindowThreadProcessId(hwnd, out processId);
                _powerPointProcessId = processId;
            }
            catch
            {
                _powerPointProcessId = 0;
            }
        }

        private void EnsurePowerPointProcessClosed()
        {
            if (_powerPointProcessId <= 0)
                return;

            try
            {
                var process = Process.GetProcessById(_powerPointProcessId);
                if (process.WaitForExit(3000))
                    return;

                process.Kill();
                process.WaitForExit(3000);
            }
            catch
            {
            }
        }

        private static void InvokeComMethod(object value, string methodName)
        {
            if (value == null)
                return;

            try
            {
                value.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, value, null);
            }
            catch
            {
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value == null)
                return;

            try
            {
                if (Marshal.IsComObject(value))
                    Marshal.FinalReleaseComObject(value);
            }
            catch
            {
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }
}
