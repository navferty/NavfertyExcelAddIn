using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.LimitTextLength
{
	internal static class TextTrimmer
	{
		public static string? LimitLength(this string value, int len, bool trimSpacedStrings)
		{
			if (string.IsNullOrEmpty(value)) return string.Empty;
			if (trimSpacedStrings && string.IsNullOrWhiteSpace(value)) return string.Empty;

			return (value.Length <= len)
				? value
				: value.Substring(0, len);
		}

	}
}
