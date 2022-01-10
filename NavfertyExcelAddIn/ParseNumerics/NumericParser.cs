using System;

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

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
			bool autoCalcEnabled = false;
			try { autoCalcEnabled = selection.Worksheet.EnableCalculation; } catch { }
=======
			var autoCalcEnabled = selection.Worksheet.EnableCalculation;
>>>>>>> ad4a2ca (iss_29 v3)
=======
			var autoCalcEnabled = selection.Worksheet.EnableCalculation;
>>>>>>> 28f09b2 (iss_29 v3)
=======
			bool autoCalcEnabled = false;
			try { autoCalcEnabled = selection.Worksheet.EnableCalculation; } catch { }
>>>>>>> 7fe5d79 (Added tests)
			if (autoCalcEnabled) selection.Worksheet.EnableCalculation = false;
			try
			{

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
				selection.ApplyForEachCellOfType2<string, object>(
					(value, cell) =>
					 {
						 var pdResult = value.ParseDecimal();
						 if (null == pdResult || !pdResult.ConvertedValue.HasValue)
=======
=======
>>>>>>> 28f09b2 (iss_29 v3)
				selection.ApplyForEachCellOfType<string, object>(
					(value, cell) =>
					 {
						 var pdResult = value.ParseDecimal();
						 if (!pdResult.ConvertedValue.HasValue)
<<<<<<< HEAD
>>>>>>> ad4a2ca (iss_29 v3)
=======
>>>>>>> 28f09b2 (iss_29 v3)
=======
				selection.ApplyForEachCellOfType2<string, object>(
					(value, cell) =>
					 {
						 var pdResult = value.ParseDecimal();
						 if (null == pdResult || !pdResult.ConvertedValue.HasValue)
>>>>>>> 7fe5d79 (Added tests)
							 return (object)value;

						 //Parsed Ok...
						 if (pdResult.IsMoney)
						 {
							 string currencyFormat = pdResult.IsCurrencyFromRU() ? EXCEL_CURRENCY_FORMAT_TEMPLATE_RUS : EXCEL_CURRENCY_FORMAT_TEMPLATE_LAT;
							 string curSymFmt = @"[$" + pdResult.Currency + @"]";
							 currencyFormat = currencyFormat.Replace(CURRENCY_TEMPLATE, curSymFmt);

							 //cell.Value = Convert.ToDouble(pdResult.ConvertedValue.Value);
							 cell.NumberFormat = currencyFormat;
						 }
						 return (object)Convert.ToDouble(pdResult.ConvertedValue.Value);// Excel stores numerics as Double
					 });
			}
			finally
			{ if (autoCalcEnabled) selection.Worksheet.EnableCalculation = autoCalcEnabled; }

		}
	}
}
