using System.Diagnostics;
using System.Windows.Forms;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Commons
{
    [DebuggerStepThrough]
    public class DialogService : IDialogService
    {
        public void ShowError(string message)
        {
            MessageBox.Show(message, UIStrings.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void ShowInfo(string message)
        {
            MessageBox.Show(message, UIStrings.Info, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        public bool Ask(string message, string caption)
        {
            return MessageBox.Show(message,
                        caption,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) == DialogResult.Yes;
        }

        public void ShowVersion()
        {
            var assembly = typeof(ThisAddIn).Assembly;
            var ver = FileVersionInfo.GetVersionInfo(assembly.Location);
            ShowInfo(string.Format(UIStrings.ShowVersionMessage, ver.FileVersion));
        }
    }
}
