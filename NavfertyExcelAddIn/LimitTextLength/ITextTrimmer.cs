#nullable enable

using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.LimitTextLength
{
	public interface ITextTrimmer
	{
		void DisplayTrimTextUI(Range range);
	}
}
