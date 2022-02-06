using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Navferty.Common;

namespace NavfertyExcelAddIn.ParseNumerics
{
	/// <summary>
	/// Parser implemented in https://github.com/navferty/NumericParser
	/// </summary>
	public static class DecimalParser
	{
		//private static readonly Regex SpacesPattern = new(@"\s");
		private static readonly Regex DecimalPattern = new(@"[\d\.\,\s]+");
		private static readonly Regex ExponentPattern = new(@"[-+]?\d*\.?\d+[eE][-+]?\d+");

		public static NumericParseResult? ParseDecimal(this string value)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				return null;
			}

			var v = value.RemoveSpacesEx();

			if (ExponentPattern.IsMatch(v))
				return new NumericParseResult(v.TryParseExponent());

			if (!DecimalPattern.IsMatch(value))
				return null;//Doesn't look like a number at all.

			//Determine the decimal separator . or ,
			if (v.Contains(",") && v.Contains("."))
			{
				var last = v.LastIndexOfAny(new[] { ',', '.' });
				var c = v[last];
				return v.CountChars(c) == 1
					? v.TryParse(c == '.' ? Format.Dot : Format.Comma)
					: null;
			}

			if (v.Contains(","))
			{
				return v.CountChars(',') == 1
					? v.TryParse(Format.Comma)
					: v.TryParse(Format.Dot);
			}

			if (v.Contains("."))
			{
				return v.CountChars('.') == 1
					? v.TryParse(Format.Dot)
					: v.TryParse(Format.Comma);
			}

			return v.TryParse(Format.Dot);
		}

		private static double? TryParseExponent(this string value)
		{
			return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result)
				? result
				: null;
		}

		private static Lazy<string[]> _allCurrencySymbolsCacheLazy = new(()
			=> CultureInfo.GetCultures(CultureTypes.AllCultures)
		.Select(ci => ci.NumberFormat.CurrencySymbol)
		.Distinct()
		.Where(cur => !string.IsNullOrWhiteSpace(cur))
		.ToArray());

		private static NumericParseResult? TryParse(this string value, Format info)
		{
			var formatInfo = (NumberFormatInfo)NumberFormatInfo.InvariantInfo.Clone();

			if (info == Format.Comma)
			{
				formatInfo.CurrencyDecimalSeparator = ",";
				formatInfo.CurrencyGroupSeparator = ".";
				formatInfo.NumberDecimalSeparator = ",";
				formatInfo.NumberGroupSeparator = ".";
			}
			else
			{
				formatInfo.CurrencyDecimalSeparator = ".";
				formatInfo.CurrencyGroupSeparator = ",";
			}
			// добавить тест-кейсов на формат валют
			//formatInfo.CurrencyNegativePattern = 8;
			//formatInfo.CurrencyPositivePattern = 3;

			var valueParsed = double.TryParse(value, NumberStyles.Currency, formatInfo, out double result);
			if (valueParsed) return new NumericParseResult(result);//Parsed without our help

			//.TryParse cannot parse a string with a CurrencySymbol and number format not from the current user culture
			//Therefore, we help to parse the string with an arbitrary currency manualy...

			//detect how many different currency symbols contains source string.
			var currenciesInValue = _allCurrencySymbolsCacheLazy.Value.Where(cur => value.Contains(cur));//.ToArray();
			if (currenciesInValue.Count() == 1)// TODO: If the string contains several different currency symbols, do not convert, because currency priority is not clear
			{
				var curSymb = currenciesInValue.First();
				//Remove found currencySymbol from source string
				var valueWithoutCurrencySymbol = value.Replace(curSymb, string.Empty);
				valueParsed = double.TryParse(valueWithoutCurrencySymbol, NumberStyles.Currency, formatInfo, out result);
				if (valueParsed)
				{
					//Debug.WriteLine($"Parsed value: '{value}, valueWithoutCurrencySymbol: {valueWithoutCurrencySymbol}', result: {result}, currency: {curSymb}");
					return new NumericParseResult(result, curSymb);
				}
				//It was not possible to parse the line, even after removing the currencySymbol, most likely this is not about money at all...
			}
			return null;//Not found any currency symbols, or found more than one, or even not number...
		}

		private enum Format
		{
			Dot,
			Comma
		}
	}
}
