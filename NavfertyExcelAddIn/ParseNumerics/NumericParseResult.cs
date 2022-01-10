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

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
=======

		public NumericParseResult() { }

>>>>>>> ad4a2ca (iss_29 v3)
=======

		public NumericParseResult() { }

>>>>>>> 28f09b2 (iss_29 v3)
=======
>>>>>>> 7fe5d79 (Added tests)
		public NumericParseResult(decimal? value, string curr = "")
		{
			ConvertedValue = value;
			if (null != curr) Currency = curr.Trim();
			IsMoney = (null != Currency && !string.IsNullOrEmpty(Currency));
		}

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
		public bool IsCurrencyFromCurrentCulture()
			=> (_currencySystem == Currency);

		public bool IsCurrencyFromRU()
			=> (_currencyRu == Currency);

		/// <summary>This code is sample from internet - WAS NOT CHECKED!!!</summary>
		private static string GetCurrencyFormat(CultureInfo culture = null)
=======
=======
>>>>>>> 28f09b2 (iss_29 v3)
		public bool IsCurrencyFromCurrentCulture() => (_currencySystem == Currency);
=======
		public bool IsCurrencyFromCurrentCulture()
			=> (_currencySystem == Currency);
>>>>>>> 7fe5d79 (Added tests)

		public bool IsCurrencyFromRU()
			=> (_currencyRu == Currency);

		/// <summary>This code is sample from internet - WAS NOT CHECKED!!!</summary>
<<<<<<< HEAD
		public static string GetCurrencyFormat(CultureInfo culture = null)
<<<<<<< HEAD
>>>>>>> ad4a2ca (iss_29 v3)
=======
>>>>>>> 28f09b2 (iss_29 v3)
=======
		private static string GetCurrencyFormat(CultureInfo culture = null)
>>>>>>> 7fe5d79 (Added tests)
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

<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
		public static bool operator ==(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return true;
			if (obj1 is null) return false;

			return obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney;
		}

		public static bool operator !=(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return false;
			if (obj1 is null) return true;

			return !(obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney);
		}

		public override bool Equals(object obj)
			=> !(obj is null) && (obj.GetType() == typeof(NumericParseResult) && (this == obj as NumericParseResult));
=======

=======
>>>>>>> 0a251f3 (fixed bugs in NumericParseResult comparsion and _allCurrencySymbolsCache initialization)
		public static bool operator ==(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return true;
			if (obj1 is null) return false;

			return obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney;
		}

		public static bool operator !=(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return false;
			if (obj1 is null) return true;

			return !(obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney);
		}

		public override bool Equals(object obj)
<<<<<<< HEAD
			=> (null != obj) && (obj.GetType() == typeof(NumericParseResult) && (this == obj as NumericParseResult));
>>>>>>> 1a6c523 (fix azure-pipelines-publish.yml)
=======
>>>>>>> ad4a2ca (iss_29 v3)
=======
>>>>>>> 28f09b2 (iss_29 v3)

=======
>>>>>>> 15bb3e9 (iss_29 fixed bugs in NumericParseResult comparsion and _allCurrencySymbolsCache initialization)
		public static bool operator ==(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return true;
			if (obj1 is null) return false;

			return obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney;
		}

		public static bool operator !=(NumericParseResult obj1, NumericParseResult obj2)
		{
			if ((obj1 is null) && (obj2 is null)) return false;
			if (obj1 is null) return true;

			return !(obj1.ConvertedValue == obj2.ConvertedValue
				&& obj1.Currency == obj2.Currency
				&& obj1.IsMoney == obj2.IsMoney);
		}

		public override bool Equals(object obj)
=======
>>>>>>> 0a251f3 (fixed bugs in NumericParseResult comparsion and _allCurrencySymbolsCache initialization)
			=> !(obj is null) && (obj.GetType() == typeof(NumericParseResult) && (this == obj as NumericParseResult));



	}
}
