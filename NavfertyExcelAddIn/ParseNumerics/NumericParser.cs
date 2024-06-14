using System;
using System.Diagnostics;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.ParseNumerics
{
	public class NumericParser : INumericParser
	{
		public void Parse(Range selection)
		{
			//Формат положительных;Формат отрицательных;Формат нулей;Формат текста
			const string EXCEL_CURRENCY_FORMAT_TEMPLATE_RUS = @"_-* #,##0.00 {CUR}_-;-* #,##0.00 {CUR}_-;_-* ""-""?? {CUR}_-;_-@_-";
			const string EXCEL_CURRENCY_FORMAT_TEMPLATE_LAT = @"_-{CUR}* # ##0.00_-;-{CUR}* # ##0.00_-;_-{CUR}* ""-""??_-;_-@_-";
			const string CURRENCY_TEMPLATE = @"{CUR}";

			bool autoCalcEnabled = false;
			try { autoCalcEnabled = selection.Worksheet.EnableCalculation; } catch { }
			if (autoCalcEnabled) selection.Worksheet.EnableCalculation = false;

#if DEBUG
			var sw = new Stopwatch();
			sw.Start();
#endif
			try
			{

				selection.ApplyForEachCellOfType2<string, object>(
					(value, cell) =>
					 {
						 var pdResult = value.ParseDecimal();
						 if (!pdResult.HasValue || !pdResult.Value.ConvertedValue.HasValue)
							 return (object)value;

						 var npr = pdResult.Value;
						 //Parsed Ok...
						 if (pdResult.Value.IsMoney)
						 {
							 string currencyFormat = npr.IsCurrencyFromRU() ? EXCEL_CURRENCY_FORMAT_TEMPLATE_RUS : EXCEL_CURRENCY_FORMAT_TEMPLATE_LAT;
							 string curSymFmt = @"[$" + npr.Currency + @"]";
							 cell.NumberFormat = currencyFormat.Replace(CURRENCY_TEMPLATE, curSymFmt);
						 }
						 return (object)Convert.ToDouble(npr.ConvertedValue.Value);// Excel stores numerics as Double
					 });
			}
			finally
			{
#if DEBUG
				sw.Stop();
				Debug.WriteLine($"NumericParser.Parse() {sw.Elapsed.TotalMilliseconds}ms");
#endif

				if (autoCalcEnabled) selection.Worksheet.EnableCalculation = autoCalcEnabled;//Restart sheet autorecalc
			}
		}
	}
}
