using System;
using System.Windows.Forms;

namespace InstallerCA
{
    public class WindowWrapper : IWin32Window
    {
        private readonly IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }
}