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
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && _formats.Contains(Path.GetExtension(path).ToLowerInvariant());
        }

        public void Prepare(string path, ContextObject context)
        {
            // Match QuickLook.Plugin.OfficeViewer's preferred size (1920x1440, 0.9)
            // so that navigating between PowerPoint files and Word/Excel/other Office
            // documents does NOT resize the shared preview window. The window resize
            // is what makes QuickLook's Mica/Acrylic title bar composite lag and ghost
            // the previous file name over the new one while switching.
            context.SetPreferredSizeFit(new Size { Width = 1920, Height = 1440 }, 0.9d);
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
