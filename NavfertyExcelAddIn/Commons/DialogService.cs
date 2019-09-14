using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Excel.Application;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Commons
{
    [DebuggerStepThrough]
    public class DialogService : IDialogService
    {
		private Application App => Globals.ThisAddIn.Application;

        private static readonly Dictionary<FileType, FileExtensionFilter> ExtensionFilters
            = new Dictionary<FileType, FileExtensionFilter>
            {
                {FileType.All, new FileExtensionFilter("All files", "*.*")},
                {FileType.Excel, new FileExtensionFilter("Excel files", "*.xls; *.xlsx; *.xlsm; *.xlsb")},
                {FileType.Xml, new FileExtensionFilter("XML Files", "*.xml")},
                {FileType.Xsd, new FileExtensionFilter("XSD Files", "*.xsd")},
                {FileType.Pdf, new FileExtensionFilter("PDF files", "*.pdf")}
            };

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

        public IReadOnlyCollection<string> AskForFiles(bool allowMultiselect, FileType fileType)
        {
            var dialog = App.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];

            dialog.AllowMultiSelect = allowMultiselect;
            dialog.InitialView = MsoFileDialogView.msoFileDialogViewList;

            var fileExtension = ExtensionFilters[fileType];
            dialog.Filters.Add(fileExtension.Description, fileExtension.Extentions, 1);
            dialog.FilterIndex = 1;

            dialog.Show();

            return dialog.SelectedItems.Cast<string>().ToArray();
        }

        public string AskFileNameSaveAs(string initialFileName, FileType fileType)
        {
            // TODO can we use build-in excel dialog?
            using (var dialog = new SaveFileDialog())
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialog.FileName = initialFileName;

                var fileExtension = ExtensionFilters[fileType];
                dialog.Filter = $@"{fileExtension.Description} | {fileExtension.Extentions}";

                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.FileName
                    : string.Empty;
            }
        }
    }
}
