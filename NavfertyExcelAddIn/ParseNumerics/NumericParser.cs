using System;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

using NumericParser;

namespace NavfertyExcelAddIn.ParseNumerics;

public class NumericParser : INumericParser
{
	public void Parse(Range selection)
	{
		selection.ApplyForEachCellOfType<string, object>(
			value =>
			{
				if (value.TryParseDecimal(out var newValue))
				{
					// Excel stores numerics as Double
					return Convert.ToDouble(newValue);
				}

				return value;
			});
	}
}
