using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.ParseNumerics
{
	public struct NumericParseResult
	{
		private static readonly string currencySymbolFromOSUserLocale = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;
		private static readonly string currencySymbolRu = CultureInfo.GetCultureInfo("ru").NumberFormat.CurrencySymbol;

		public readonly double? ConvertedValue = null;
		public readonly string? Currency = null;

		public NumericParseResult(double? value, string? curr = null)
		{
			ConvertedValue = value;
			Currency = curr;
		}

		public bool IsMoney
			=> !string.IsNullOrEmpty(Currency);

		public bool IsCurrencyFromCurrentCulture()
			=> (currencySymbolFromOSUserLocale == Currency);

		public bool IsCurrencyFromRU()
			=> (currencySymbolRu == Currency);

		/*
		 
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

		*/

	}
}
