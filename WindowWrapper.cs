using System;
using System.Windows.Forms;

namespace WpfCore.FolderPicker
{
    public class WindowWrapper : IWin32Window
    {
        public IntPtr Handle { get; }
        public WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }
    }
}