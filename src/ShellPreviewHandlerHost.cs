using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Input;
using Microsoft.Win32;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    internal sealed class ShellPreviewHandlerHost : HwndHost
    {
        private const string PreviewHandlerKeyPath = @"shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}";
        private const uint StgmRead = 0x00000000;
        private const uint ClsctxInprocServer = 0x1;
        private const uint ClsctxLocalServer = 0x4;
        private const int WsChild = 0x40000000;
        private const int WsVisible = 0x10000000;
        private const uint SwpNoZOrder = 0x0004;
        private const uint SwpNoActivate = 0x0010;

        private static readonly WindowProc HostWindowProc = DefWndProc;
        private static bool _windowClassRegistered;

        private readonly string _path;
        private IntPtr _hostHwnd;
        private IPreviewHandler _previewHandler;
        private object _previewHandlerObject;
        private IStream _fileStream;
        private bool _started;
        private bool _disposed;

        public ShellPreviewHandlerHost(string path)
        {
            _path = path;
            Loaded += delegate { StartPreview(); };
        }

        public event EventHandler PreviewLoaded;

        public event EventHandler<ShellPreviewFailedEventArgs> PreviewFailed;

        public static bool CanPreview(string path)
        {
            return GetPreviewHandlerGuid(Path.GetExtension(path)) != null;
        }

        protected override HandleRef BuildWindowCore(HandleRef hwndParent)
        {
            EnsureWindowClassRegistered();

            _hostHwnd = CreateWindowEx(
                0,
                "QuickLookPowerPointShellPreviewHost",
                string.Empty,
                WsChild | WsVisible,
                0,
                0,
                Math.Max(1, (int)ActualWidth),
                Math.Max(1, (int)ActualHeight),
                hwndParent.Handle,
                IntPtr.Zero,
                IntPtr.Zero,
                IntPtr.Zero);

            if (_hostHwnd == IntPtr.Zero)
                throw new InvalidOperationException("Could not create Shell Preview Handler host window.");

            return new HandleRef(this, _hostHwnd);
        }

        protected override void DestroyWindowCore(HandleRef hwnd)
        {
            ClearPreview();

            if (hwnd.Handle != IntPtr.Zero)
                DestroyWindow(hwnd.Handle);

            _hostHwnd = IntPtr.Zero;
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                ClearPreview();
            }

            base.Dispose(disposing);
        }

        protected override void OnWindowPositionChanged(Rect rcBoundingBox)
        {
            base.OnWindowPositionChanged(rcBoundingBox);
            UpdatePreviewRect((int)Math.Max(1, rcBoundingBox.Width), (int)Math.Max(1, rcBoundingBox.Height));
        }

        private void StartPreview()
        {
            if (_started || _disposed)
                return;

            _started = true;

            try
            {
                var handlerGuid = GetPreviewHandlerGuid(Path.GetExtension(_path));
                if (handlerGuid == null)
                    throw new InvalidOperationException("No Shell Preview Handler is registered for this file type.");

                _previewHandlerObject = CreatePreviewHandler(handlerGuid.Value);
                _previewHandler = _previewHandlerObject as IPreviewHandler;
                if (_previewHandler == null)
                    throw new InvalidOperationException("The registered Shell Preview Handler does not expose IPreviewHandler.");

                InitializePreviewHandler(_previewHandlerObject);

                var rect = CreateRect(Math.Max(1, (int)ActualWidth), Math.Max(1, (int)ActualHeight));
                _previewHandler.SetWindow(_hostHwnd, ref rect);
                _previewHandler.DoPreview();
                UpdatePreviewRect(rect.Right, rect.Bottom);
                FocusPreviewHandler();

                var loaded = PreviewLoaded;
                if (loaded != null)
                    loaded(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                ClearPreview();

                var failed = PreviewFailed;
                if (failed != null)
                    failed(this, new ShellPreviewFailedEventArgs(ex));
            }
        }

        protected override bool TabIntoCore(TraversalRequest request)
        {
            FocusPreviewHandler();
            return true;
        }

        protected override void OnGotKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            base.OnGotKeyboardFocus(e);
            FocusPreviewHandler();
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            FocusPreviewHandler();
        }

        protected override bool TranslateAcceleratorCore(ref MSG msg, ModifierKeys modifiers)
        {
            if (_previewHandler == null)
                return base.TranslateAcceleratorCore(ref msg, modifiers);

            var nativeMsg = new NativeMsg
            {
                Hwnd = msg.hwnd,
                Message = msg.message,
                WParam = msg.wParam,
                LParam = msg.lParam,
                Time = msg.time,
                PtX = msg.pt_x,
                PtY = msg.pt_y
            };

            var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NativeMsg)));
            try
            {
                Marshal.StructureToPtr(nativeMsg, ptr, false);
                return _previewHandler.TranslateAccelerator(ptr) == 0;
            }
            catch
            {
                return base.TranslateAcceleratorCore(ref msg, modifiers);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }

        private void FocusPreviewHandler()
        {
            if (_previewHandler == null)
                return;

            try
            {
                _previewHandler.SetFocus();
            }
            catch
            {
            }
        }

        private void InitializePreviewHandler(object handler)
        {
            var initWithFile = handler as IInitializeWithFile;
            if (initWithFile != null)
            {
                initWithFile.Initialize(_path, StgmRead);
                return;
            }

            var initWithStream = handler as IInitializeWithStream;
            if (initWithStream != null)
            {
                SHCreateStreamOnFileEx(_path, StgmRead, 0, false, null, out _fileStream);
                initWithStream.Initialize(_fileStream, StgmRead);
                return;
            }

            var initWithItem = handler as IInitializeWithItem;
            if (initWithItem != null)
            {
                var iid = typeof(IShellItem).GUID;
                IShellItem shellItem;
                var hr = SHCreateItemFromParsingName(_path, IntPtr.Zero, ref iid, out shellItem);
                Marshal.ThrowExceptionForHR(hr);
                initWithItem.Initialize(shellItem, StgmRead);
                return;
            }

            throw new InvalidOperationException("The Shell Preview Handler cannot be initialized with a file, stream, or shell item.");
        }

        private void UpdatePreviewRect(int width, int height)
        {
            if (_hostHwnd != IntPtr.Zero)
                SetWindowPos(_hostHwnd, IntPtr.Zero, 0, 0, width, height, SwpNoZOrder | SwpNoActivate);

            if (_previewHandler != null)
            {
                try
                {
                    var rect = CreateRect(width, height);
                    _previewHandler.SetRect(ref rect);
                }
                catch
                {
                }
            }
        }

        private void ClearPreview()
        {
            if (_previewHandler != null)
            {
                try
                {
                    _previewHandler.Unload();
                }
                catch
                {
                }
            }

            ReleaseComObject(_previewHandler);
            if (!ReferenceEquals(_previewHandler, _previewHandlerObject))
                ReleaseComObject(_previewHandlerObject);

            _previewHandler = null;
            _previewHandlerObject = null;

            ReleaseComObject(_fileStream);
            _fileStream = null;
        }

        private static object CreatePreviewHandler(Guid clsid)
        {
            var iid = typeof(IPreviewHandler).GUID;
            object instance;
            var hr = CoCreateInstance(ref clsid, IntPtr.Zero, ClsctxInprocServer | ClsctxLocalServer, ref iid, out instance);
            Marshal.ThrowExceptionForHR(hr);
            return instance;
        }

        private static Guid? GetPreviewHandlerGuid(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                return null;

            var guidText = GetPreviewHandlerGuidText(extension);
            Guid guid;
            return Guid.TryParse(guidText, out guid) ? guid : (Guid?)null;
        }

        private static string GetPreviewHandlerGuidText(string extension)
        {
            using (var extensionKey = Registry.ClassesRoot.OpenSubKey(extension))
            using (var extensionPreviewKey = extensionKey == null ? null : extensionKey.OpenSubKey(PreviewHandlerKeyPath))
            {
                var directGuid = extensionPreviewKey == null ? null : extensionPreviewKey.GetValue(null) as string;
                if (!string.IsNullOrEmpty(directGuid))
                    return directGuid;

                var className = extensionKey == null ? null : extensionKey.GetValue(null) as string;
                if (string.IsNullOrEmpty(className))
                    return null;

                using (var classKey = Registry.ClassesRoot.OpenSubKey(className))
                using (var classPreviewKey = classKey == null ? null : classKey.OpenSubKey(PreviewHandlerKeyPath))
                {
                    return classPreviewKey == null ? null : classPreviewKey.GetValue(null) as string;
                }
            }
        }

        private static RectNative CreateRect(int width, int height)
        {
            return new RectNative
            {
                Left = 0,
                Top = 0,
                Right = width,
                Bottom = height
            };
        }

        private static void EnsureWindowClassRegistered()
        {
            if (_windowClassRegistered)
                return;

            var wndClass = new WindowClass
            {
                Style = 0,
                WndProc = HostWindowProc,
                ClassExtra = 0,
                WindowExtra = 0,
                Instance = IntPtr.Zero,
                Icon = IntPtr.Zero,
                Cursor = IntPtr.Zero,
                Background = IntPtr.Zero,
                MenuName = null,
                ClassName = "QuickLookPowerPointShellPreviewHost"
            };

            var atom = RegisterClass(ref wndClass);
            if (atom == 0)
            {
                var error = Marshal.GetLastWin32Error();
                if (error != 1410)
                    throw new InvalidOperationException("Could not register Shell Preview Handler host window class. Win32 error: " + error);
            }

            _windowClassRegistered = true;
        }

        private static IntPtr DefWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            return DefWindowProc(hwnd, msg, wParam, lParam);
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

        [DllImport("ole32.dll", PreserveSig = true)]
        private static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string path, IntPtr bindingContext, ref Guid riid, out IShellItem shellItem);

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateStreamOnFileEx([MarshalAs(UnmanagedType.LPWStr)] string fileName, uint mode, uint attributes, bool create, IStream template, out IStream stream);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern ushort RegisterClass(ref WindowClass lpWndClass);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateWindowEx(int exStyle, string className, string windowName, int style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyWindow(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("user32.dll")]
        private static extern IntPtr DefWindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        private delegate IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WindowClass
        {
            public uint Style;
            public WindowProc WndProc;
            public int ClassExtra;
            public int WindowExtra;
            public IntPtr Instance;
            public IntPtr Icon;
            public IntPtr Cursor;
            public IntPtr Background;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string MenuName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string ClassName;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RectNative
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeMsg
        {
            public IntPtr Hwnd;
            public int Message;
            public IntPtr WParam;
            public IntPtr LParam;
            public int Time;
            public int PtX;
            public int PtY;
        }

        [ComImport]
        [Guid("8895b1c6-b41f-4c1c-a562-0d564250836f")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPreviewHandler
        {
            void SetWindow(IntPtr hwnd, ref RectNative rect);

            void SetRect(ref RectNative rect);

            void DoPreview();

            void Unload();

            void SetFocus();

            void QueryFocus(out IntPtr hwnd);

            [PreserveSig]
            int TranslateAccelerator(IntPtr msg);
        }

        [ComImport]
        [Guid("b7d14566-0509-4cce-a71f-0a554233bd9b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IInitializeWithFile
        {
            void Initialize([MarshalAs(UnmanagedType.LPWStr)] string filePath, uint mode);
        }

        [ComImport]
        [Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IInitializeWithStream
        {
            void Initialize(IStream stream, uint mode);
        }

        [ComImport]
        [Guid("7f73be3f-fb79-493c-a6c7-7ee14e245841")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IInitializeWithItem
        {
            void Initialize(IShellItem shellItem, uint mode);
        }

        [ComImport]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
        }

        internal sealed class ShellPreviewFailedEventArgs : EventArgs
        {
            public ShellPreviewFailedEventArgs(Exception exception)
            {
                Exception = exception;
            }

            public Exception Exception { get; private set; }
        }
    }
}
