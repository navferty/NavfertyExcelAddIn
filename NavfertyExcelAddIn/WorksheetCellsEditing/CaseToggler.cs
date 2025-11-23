using System;
using System.Globalization;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

using NLog;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public class CaseToggler : ICaseToggler
	{
		private readonly TextInfo textInfo = new CultureInfo("en-US").TextInfo;
		private readonly ILogger logger = LogManager.GetCurrentClassLogger();

		public void ToggleCase(Range range)
		{
			var cells = range.Cast<Range>();

			var firstValue = cells
				.FirstOrDefault(CheckForLetters)
				?.Value
				.ToString();

			if (string.IsNullOrEmpty(firstValue))
			{
				return;
			}

			var targertCase = DetectTargetCase((string)firstValue!);

			range.ApplyForEachCellOfType<string, string?>(value => ConvertTextCase(value, targertCase));
		}

		private static bool CheckForLetters(Range cell)
		{
			return cell.Value is string value
				   && !string.IsNullOrWhiteSpace(value)
				   && value.Any(char.IsLetter);
		}

		private TextCaseType DetectTargetCase(string value)
		{
			logger.Debug($"Detecting case of value {value}");

			// TODO case of value 'a' - short cycle A->a->A->a
			// if(value.Length==1 && char.IsLetter(value[0])){ return ...}

			if (value.All(c => char.IsLower(c) || !char.IsLetter(c)))
			{
				// transform lower to UPPER
				logger.Debug("Detected that each char is lowercase or not letter");
				return TextCaseType.Upper;
			}
			if (value.All(c => char.IsUpper(c) || !char.IsLetter(c)))
			{
				// transform UPPER to Capital
				logger.Debug("Detected that each char is uppercase or not letter");
				return CountChars(value) == 1
					? TextCaseType.Lower
					: TextCaseType.Capital;
			}
			// transform to lower
			logger.Debug("Target case is lower (neither all chars are upper nor lower)");
			return TextCaseType.Lower;
		}

		private string ConvertTextCase(string value, TextCaseType targertCase)
		{
			switch (targertCase)
			{
				case TextCaseType.Upper:
					{
						return value.ToUpper(CultureInfo.CurrentCulture);
					}
				case TextCaseType.Lower:
					{
						return value.ToLower(CultureInfo.CurrentCulture);
					}
				case TextCaseType.Capital:
					{
						// TextInfo.ToTitleCase ignores words with all capital chars
						var lower = value.ToLower(CultureInfo.CurrentCulture);
						return textInfo.ToTitleCase(lower);
					}
				default:
					{
						throw new ArgumentOutOfRangeException(nameof(targertCase), targertCase, @"WTF is the text case???");
					}
			}
		}

		private static int CountChars(string value) => value.Count(char.IsLetter);

		private enum TextCaseType
		{
			Upper,
			Lower,
			Capital
		}
	}
}
