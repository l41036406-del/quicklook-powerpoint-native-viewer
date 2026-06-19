using System.IO;
using System.Linq;
using System.Windows;
using QuickLook.Common.Plugin;

namespace QuickLook.Plugin.PowerPointNativeViewer
{
    public class Plugin : IViewer
    {
        private readonly string[] _formats = { ".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".pps", ".ppsx", ".ppsm" };

        public int Priority
        {
            get { return 100; }
        }

        public void Init()
        {
            PowerPointImagePreviewControl.CleanupOldSessions();
            PowerPointPdfPreviewControl.CleanupOldSessions();
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && _formats.Contains(Path.GetExtension(path).ToLowerInvariant());
        }

        public void Prepare(string path, ContextObject context)
        {
            context.SetPreferredSizeFit(new Size { Width = 1280, Height = 800 }, 0.9d);
            context.CanResize = true;
        }

        public void View(string path, ContextObject context)
        {
            var viewer = new PowerPointImagePreviewControl(path);

            context.ViewerContent = viewer;
            context.Title = Path.GetFileName(path);
            context.IsBusy = false;
        }

        public void Cleanup()
        {
        }
    }
}
