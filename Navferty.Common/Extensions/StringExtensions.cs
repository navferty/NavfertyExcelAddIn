using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

#nullable enable

namespace Navferty.Common
{
	public static class StringExtensions
	{
		//private static readonly ILogger logger = LogManager.GetCurrentClassLogger();
		private static readonly Regex spacesRegex = new Regex("\\s+", RegexOptions.None);

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static string? TrimSpaces(this string? value)
		{
			if (string.IsNullOrWhiteSpace(value))
				return null;

			// replace any single or multiple space chars with single space
			var newValue = spacesRegex.Replace(value!, " ");

			newValue = string.IsNullOrEmpty(newValue)
				? null
				: newValue.Trim();

			return newValue;
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static int CountChars(this string value, char c)
		{
			return value.Count(x => x == c);
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static string LimitLength(
			this string text,
			int MaxLength)
		{
			if (text.Length > MaxLength) text = new string(text.Take(MaxLength).ToArray());
			return text;
		}
	}
}
