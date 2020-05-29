using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.ParseNumerics
{
	public interface INumericParser
	{
		void Parse(Range selection);
	}
}
