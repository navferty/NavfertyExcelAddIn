using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
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
			=> value.Count(x => x == c);

		/// <summary>Removes only ASCII Space char (0x32)</summary>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static string RemoveSpacesFast(this string source)
			=> source.Replace(" ", string.Empty);

		/// <summary>Removes all Unicode character which is categorized as white space.</summary>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static string RemoveSpacesEx(this string source)
			=> string.Concat(source.Where(c => !char.IsWhiteSpace(c)));
	}
}
