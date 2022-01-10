using System.Globalization;
using System.Linq;
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

		//static System.Collections.Generic.Dictionary<string, System.Globalization.CultureInfo> _dicCultures = null;
		static string[] _allCurrencySymbols = null;
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

			var valueParsed = decimal.TryParse(value, NumberStyles.Currency, formatInfo, out decimal result);
			if (valueParsed) return new NumericParseResult(result);

			//decimal.TryParse не может разобрать строку со значком любой валюты, кроме валюты текущей культуры,
			//и символ валюты должен располагаться в правильном месте (как требуется в конкретной культуре)!!!

			if (_allCurrencySymbols == null)//Fill once static currency symbols list
			{
				_allCurrencySymbols = (from ci in System.Globalization.CultureInfo.GetCultures(CultureTypes.AllCultures)
									   let curSymb = ci.NumberFormat.CurrencySymbol
									   where (null != curSymb && !string.IsNullOrWhiteSpace(curSymb.Trim()))
									   orderby curSymb ascending
									   select curSymb).Distinct().ToArray();

				/*			
				Вообще, надо бы хранить Dictionary<CurrencySymbols/CultureInfo>, 
				чтобы выбрать по коду валюты соответствующую культуру и по ней строить форматирование валютной страки...
				но у нескольких РАЗНЫХ культур часто одинаковые знаки валюты, и хз какую использовать...

				_dicCultures = new System.Collections.Generic.Dictionary<string, System.Globalization.CultureInfo>();
				var cultures = (from cult in System.Globalization.CultureInfo.GetCultures(CultureTypes.NeutralCultures)
								where ((cult != null)
								&& (cult.Parent == System.Globalization.CultureInfo.InvariantCulture)
								&& (null != cult.NumberFormat.CurrencySymbol)
								&& (!string.IsNullOrWhiteSpace(cult.NumberFormat.CurrencySymbol.Trim())))
								orderby cult.NumberFormat.CurrencySymbol ascending, cult.EnglishName ascending
								select cult)
								.GroupBy(ci => ci.NumberFormat.CurrencySymbol)
								.Select(g => g.First()).ToArray();

				cultures.ToList().ForEach(cult => _dicCultures.Add(cult.NumberFormat.CurrencySymbol, cult));
				*/
			}

			//detect how many currency symbols contains our source string...
			var currenciesInValue = _allCurrencySymbols.Where(cur => value.Contains(cur));
			if (currenciesInValue.Count() >= 1)// TODO: Если строка содержит несколько разных символов валют надо что-то сделать... 
			{
				var curSymb = currenciesInValue.First();
				//Remove found currency from source string
				var valueWithoutCurrencySymbol = value.Replace(curSymb, string.Empty);
				valueParsed = decimal.TryParse(valueWithoutCurrencySymbol, NumberStyles.Currency, formatInfo, out result);
				if (!valueParsed)
				{
					return new NumericParseResult();
				}

				//System.Windows.Forms.MessageBox.Show($"Parsed value: '{value}, valueWithoutCurrencySymbol: {valueWithoutCurrencySymbol}', result: {result}, currency: {curSymb}");
				return new NumericParseResult(result, curSymb);
			}

			//Not found any currency symbols, or found more than one!
			return new NumericParseResult();
		}

		private enum Format
		{
			Dot,
			Comma
		}
	}
}
