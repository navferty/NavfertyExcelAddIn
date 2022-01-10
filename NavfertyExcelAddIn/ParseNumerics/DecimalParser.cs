using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace NavfertyExcelAddIn.ParseNumerics
{
	/// <summary>
	/// Parser implemented in https://github.com/navferty/NumericParser
	/// </summary>
	public static class DecimalParser
	{
		private static readonly Regex SpacesPattern = new Regex(@"\s");
		private static readonly Regex DecimalPattern = new Regex(@"[\d\.\,\s]+");
		private static readonly Regex ExponentPattern = new Regex(@"[-+]?\d*\.?\d+[eE][-+]?\d+");

		public static NumericParseResult ParseDecimal(this string value)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				return new NumericParseResult();
			}

			var v = SpacesPattern.Replace(value, match => string.Empty);

			if (ExponentPattern.IsMatch(v))
			{
				return new NumericParseResult(v.TryParseExponent());
			}

			if (!DecimalPattern.IsMatch(value))
			{
				return new NumericParseResult();
			}

			if (v.Contains(",") && v.Contains("."))
			{
				var last = v.LastIndexOfAny(new[] { ',', '.' });
				var c = v[last];
				return v.CountChars(c) == 1
					? v.TryParse(c == '.' ? Format.Dot : Format.Comma)
					: new NumericParseResult();
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

		private static int CountChars(this string value, char c)
		{
			return value.Count(x => x == c);
		}

		private static decimal? TryParseExponent(this string value)
		{
			return decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal result)
				? result
				: (decimal?)null;
		}

		private static string[] _allCurrencySymbolsCache = null;
		/// <summary>Chache for all known currency symbols from Globalization</summary>
		private static string[] GetAllCurrencySymbols()
		{
			//Fill once on first query
			_allCurrencySymbolsCache = _allCurrencySymbolsCache
				?? (from ci in CultureInfo.GetCultures(CultureTypes.AllCultures)
					let curSymb = ci.NumberFormat.CurrencySymbol
					where (null != curSymb && !string.IsNullOrWhiteSpace(curSymb.Trim()))
					orderby curSymb ascending
					select curSymb).Distinct().ToArray();

			/* 
			 //code for get currency/countries relation for excel 'Cultures.xlsx'
			var culturesByCurrency = CultureInfo.GetCultures(CultureTypes.AllCultures)
				.Where(ci => ci.Parent == CultureInfo.InvariantCulture)
				.Select(ci => new { ci, ci.NumberFormat.CurrencySymbol })
				.Where(x => !string.IsNullOrWhiteSpace(x.CurrencySymbol))
				.GroupBy(x => x.CurrencySymbol)
				.Select(x => new { Symbol = x.First().CurrencySymbol, Cultures = x.Select(z => z.ci).ToArray() })
				.ToDictionary(x => x.Symbol, x => x.Cultures);


			StringBuilder sb = new StringBuilder();
			sb.AppendLine("Currency\tCount\tCultures");
			var yyy = (from kvp in culturesByCurrency
					   orderby kvp.Value.Length descending
					   select new { Currency = kvp.Key, CultureCount = kvp.Value.Length, aCultures = kvp.Value }).ToArray();

			yyy.ToList().ForEach(x =>
			{
				var aCultures = (from ci in x.aCultures
								 let name = ci.NativeName
								 orderby name ascending
								 select $"{name} ({ci.Name})").ToArray();

				var sCultures = string.Join(", ", aCultures);
				sb.AppendLine($"{x.Currency}\t{x.CultureCount}\t{sCultures}");
			});			 

			var ss = sb.ToString();
			Debug.WriteLine(ss);
			*/

			return _allCurrencySymbolsCache;
		}

		private static NumericParseResult TryParse(this string value, Format info)
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
			return decimal.TryParse(value, NumberStyles.Currency, formatInfo, out decimal result)
				? result
				: (decimal?)null;
		}

		private enum Format
		{
			Dot,
			Comma
		}
	}
}
