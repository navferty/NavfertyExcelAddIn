using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
	internal class WorksheetRow
	{
		public readonly Worksheet Sheet;

		public WorksheetRow(Worksheet ws)
		{
			Sheet = ws;
			//Displayname = ws.Name, 

			//var bIsProtected = Sheet.ProtectionMode;
		}

		public override string ToString()
		{
			var sProtection = Sheet.ProtectionMode ? " (лист ужё защищён)" : "";
			return $"{Sheet.Name}{sProtection}";
		}
	}
}
