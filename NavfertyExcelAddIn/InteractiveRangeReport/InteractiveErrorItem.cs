using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.InteractiveRangeReport
{
	public class InteractiveErrorItem
	{
		public Range? Range { get; set; }

		public string ErrorMessage { get; set; } = string.Empty;
		public string Value { get; set; } = string.Empty;
		public string Address { get; set; } = string.Empty;
		public string WorksheetName { get; set; } = string.Empty;
	}
}
