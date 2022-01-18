using System.Linq;
using System.Text.RegularExpressions;

using NLog;

namespace NavfertyExcelAddIn.Commons
{
	public static class StringExtensions
	{
		private static readonly ILogger logger = LogManager.GetCurrentClassLogger();
		private static readonly Regex spacesRegex = new Regex("\\s+", RegexOptions.None);

		public static string TrimSpaces(this string value)
		{
			if (string.IsNullOrWhiteSpace(value))
				return null;

			// replace any single or multiple space chars with single space
			var newValue = spacesRegex.Replace(value, " ");

			newValue = string.IsNullOrEmpty(newValue)
				? null
				: newValue.Trim();

			return newValue;
		}


		public static int CountChars(this string value, char c)
		{
			return value.Count(x => x == c);
		}
	}
}
