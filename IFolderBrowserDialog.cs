using System.Windows.Forms;

namespace WpfCore.FolderPicker
{
    public interface IFolderBrowserDialog
    {
        string InitialFolder { get; set; }
        string DefaultFolder { get; set; }
        string Folder { get; set; }
        DialogResult ShowDialog();
        DialogResult ShowDialog(IWin32Window owner);
        DialogResult ShowVistaDialog(IWin32Window owner);
        DialogResult ShowLegacyDialog(IWin32Window owner);
        void Dispose();
    }
}