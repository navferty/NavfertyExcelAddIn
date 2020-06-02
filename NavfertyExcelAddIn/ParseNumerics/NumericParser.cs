using System;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.ParseNumerics
{
	public class NumericParser : INumericParser
	{
		public void Parse(Range selection)
		{
			selection.ApplyForEachCellOfType<string, object>(
				value =>
				{
					var newValue = value.ParseDecimal();
					if (newValue.HasValue)
						// Excel stores numerics as Double
						return (object)Convert.ToDouble(newValue);
					return (object)value;
				});
		}
	}
}
