using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.InteractiveRangeReport;

namespace NavfertyExcelAddIn.DataValidation
{
	public interface ICellsValueValidator
	{
		IReadOnlyCollection<InteractiveErrorItem> Validate(Range range, ValidationType validationType);
	}
}