using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WpfCore.FolderPicker
{
    public sealed class FolderBrowserDialog : IFolderBrowserDialog
    {
        public string Folder { get; set; }
        public string InitialFolder { get; set; }
        public string DefaultFolder { get; set; }
        public DialogResult ShowDialog()
        {
            return ShowDialog(owner: new WindowWrapper(IntPtr.Zero));
        }
        public DialogResult ShowDialog(IWin32Window owner)
        {
            if (Environment.OSVersion.Version.Major >= 6)
                return ShowVistaDialog(owner);
            else
                return ShowLegacyDialog(owner);
        }
        public DialogResult ShowVistaDialog(IWin32Window owner)
        {
            var frm = (NativeMethods.IFileDialog)new NativeMethods.FileOpenDialogRCW();
            frm.GetOptions(out uint options);
            options |= NativeMethods.FOS_PICKFOLDERS | NativeMethods.FOS_FORCEFILESYSTEM | NativeMethods.FOS_NOVALIDATE | NativeMethods.FOS_NOTESTFILECREATE | NativeMethods.FOS_DONTADDTORECENT;
            frm.SetOptions(options);
            if (InitialFolder != null)
            {
                var riid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
                if (NativeMethods.SHCreateItemFromParsingName(InitialFolder, IntPtr.Zero, ref riid, out NativeMethods.IShellItem directoryShellItem) == NativeMethods.S_OK)
                    frm.SetFolder(directoryShellItem);
            }
            if (DefaultFolder != null)
            {
                var riid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
                if (NativeMethods.SHCreateItemFromParsingName(this.DefaultFolder, IntPtr.Zero, ref riid, out NativeMethods.IShellItem directoryShellItem) == NativeMethods.S_OK)
                    frm.SetDefaultFolder(directoryShellItem);
            }
            if (frm.Show(owner.Handle) == NativeMethods.S_OK)
            {
                if (frm.GetResult(out NativeMethods.IShellItem shellItem) == NativeMethods.S_OK)
                {
                    if (shellItem.GetDisplayName(NativeMethods.SIGDN_FILESYSPATH, out IntPtr pszString) == NativeMethods.S_OK)
                    {
                        if (pszString != IntPtr.Zero)
                        {
                            try
                            {
                                Folder = Marshal.PtrToStringAuto(pszString);
                                return DialogResult.OK;
                            }
                            finally
                            {
                                Marshal.FreeCoTaskMem(pszString);
                            }
                        }
                    }
                }
            }
            return DialogResult.Cancel;
        }
        public DialogResult ShowLegacyDialog(IWin32Window owner)
        {
            using var frm = new SaveFileDialog();
            frm.CheckFileExists = false;
            frm.CheckPathExists = true;
            frm.CreatePrompt = false;
            frm.Filter = "|" + Guid.Empty.ToString();
            frm.FileName = "any";
            if (InitialFolder != null) { frm.InitialDirectory = InitialFolder; }
            frm.OverwritePrompt = false;
            frm.Title = "Select Folder";
            frm.ValidateNames = false;
            if (frm.ShowDialog(owner) == DialogResult.OK)
            {
                Folder = Path.GetDirectoryName(frm.FileName);
                return DialogResult.OK;
            }
            else
            {
                return DialogResult.Cancel;
            }
        }
        public void Dispose() { }
    }
}