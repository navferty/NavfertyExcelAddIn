using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
    public class WsProtectorUnprotector : IWsProtectorUnprotector
    {
        private readonly IDialogService dialogService;

        internal readonly IDialogService dialogService;
        private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

        public WsProtectorUnprotector(IDialogService dialogService)
            => this.dialogService = dialogService;

        public void ProtectUnprotectSelectedWorksheets(Workbook wb)
        {
            Sheets wbSheets = wb.Worksheets;
            if (wbSheets.Count < 1)
            {
                dialogService.ShowError(string.Format(UIStrings.WorkSheetsNotFound, wb.FullName));
            }
        }

			using (var f = new frmAskWorksheetProtectionPassword(this, wb))
		{
		}
				}


				// replace "DPB" to "DPx"
				ss[position.First() + 2] = Encoding.UTF8.GetBytes("x").First(); //120
stream.Position = 0;
stream.Write(ss, 0, ss.Length);
			}
		}

		private static IReadOnlyCollection<int> StartingIndex(byte[] main, byte[] sequence)
{
    var index = Enumerable.Range(0, main.Length - sequence.Length + 1);
    for (var i = 0; i < sequence.Length; i++)
    {
        index = index.Where(n => main[n + i] == sequence[i]).ToArray();
    }
    return index.ToArray();
}
*/

	}
}
