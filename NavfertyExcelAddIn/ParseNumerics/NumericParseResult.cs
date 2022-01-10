using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.ParseNumerics
{
	public class NumericParseResult
	{
		private static readonly string _currencySystem = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;
		private static readonly string _currencyRu = CultureInfo.GetCultureInfo("ru-RU").NumberFormat.CurrencySymbol;

		public readonly decimal? ConvertedValue = null;
		public readonly string Currency = "";
		public readonly bool IsMoney = false;


		public NumericParseResult() { }

		public NumericParseResult(decimal? value, string curr = "")
		{
			ConvertedValue = value;
			if (null != curr) Currency = curr.Trim();
			IsMoney = (null != Currency && !string.IsNullOrEmpty(Currency));
		}

		public bool IsCurrencyFromCurrentCulture() => (_currencySystem == Currency);

		public bool IsCurrencyFromRU() => (_currencyRu == Currency);

		/// <summary>This code is sample from internet - WAS NOT CHECKED!!!</summary>
		public static string GetCurrencyFormat(CultureInfo culture = null)
		{
			if (culture == null) culture = CultureInfo.CurrentCulture;

			//we'll use string.Format later to replace {0} with the currency symbol
			//and {1} with the number format
			string[] negativePatternStrings = {
				"({0}{1})",
				"-{0}{1}",
				"{0}-{1}",
				"{0}{1}-",
				"({1}{0})",
				"-{1}{0}",
				"{1}-{0}",
				"{1}{0}-",
				"-{1} {0}",
				"-{0} {1}",
				"{1} {0}-",
				"{0} {1}-",
				"{0} -{1}",
				"{1}- {0}",
				"({0} {1})",
				"({1} {0})" };
			string[] positivePatternStrings = {
				"{0}{1}",
				"{1}{0}",
				"{0} {1}",
				"{1}{0}" };

			var numberFormat = culture.NumberFormat;

			//Generate 0's to fill in the format after the decimal place
			var decimalPlaces = new string('0', numberFormat.CurrencyDecimalDigits);

			//concatenate the full number format, e.g. #,0.00
			var fullDigitFormat = $"#{numberFormat.CurrencyGroupSeparator}0{numberFormat.CurrencyDecimalSeparator}{decimalPlaces}";

			//use string.Format on the patterns to get the positive and 
			//negative formats
			var positiveFormat = string.Format(positivePatternStrings[numberFormat.CurrencyPositivePattern],
											   numberFormat.CurrencySymbol, fullDigitFormat);

			var negativeFormat = string.Format(negativePatternStrings[numberFormat.CurrencyNegativePattern],
											   numberFormat.CurrencySymbol, fullDigitFormat);

			//finally, return the full format
			return $"{positiveFormat};{negativeFormat}";
		}




	}
}
