using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyCommon;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.InteractiveRangeReport;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
	public class ErrorFinder : IErrorFinder
	{
		public IReadOnlyCollection<InteractiveErrorItem> GetAllErrorCells(Range range)
		{
			return GetErroredCells(range).ToArray();
		}

		private IEnumerable<InteractiveErrorItem> GetErroredCells(Range range)
		{
			object rangeValue = range.Value;
			var worksheetName = range.Worksheet.Name;

			if (rangeValue == null)
				yield break;

			// only single cell exists on sheet
			if (!(rangeValue is object[,] values))
			{
				if (IsXlCvErr(rangeValue))
				{
					yield return new InteractiveErrorItem
					{
						Range = range,
						Value = range.GetFormula(),
						ErrorMessage = ((CVErrEnum)rangeValue).GetEnumDescription(),
						Address = range.GetRelativeAddress(),
						WorksheetName = worksheetName
					};
				}
				yield break;
			}

			int upperI = values.GetUpperBound(0); // Columns
			int upperJ = values.GetUpperBound(1); // Rows

			for (int i = values.GetLowerBound(0); i <= upperI; i++)
			{
				for (int j = values.GetLowerBound(1); j <= upperJ; j++)
				{
					var value = values[i, j];
					if (IsXlCvErr(value))
					{
						var currentRange = (Range)range.Cells[i, j];
						yield return new InteractiveErrorItem
						{
							Range = currentRange,
							Value = currentRange.GetFormula(),
							ErrorMessage = ((CVErrEnum)value).GetEnumDescription(),
							Address = currentRange.GetRelativeAddress(),
							WorksheetName = worksheetName
						};
					}
				}
			}
		}

		private bool IsXlCvErr(object obj)
		{
			// only CVErr values in Excel are interpreted as Int32 values in .NET
			// https://stackoverflow.com/questions/16217350/vba-looking-for-error-values-in-a-specific-column
			// https://xldennis.wordpress.com/2006/11/22/dealing-with-cverr-values-in-net-%E2%80%93-part-i-the-problem/
			return (obj) is Int32;
		}
	}
}
