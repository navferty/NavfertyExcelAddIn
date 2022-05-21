#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.LimitTextLength
{
	internal class TextTrimmer : ITextTrimmer
	{
		public void DisplayTrimTextUI(Range range)
		{
			int maxLen = range
				.Cells?
				.Cast<Range>()?
				.Select(cell => (string.IsNullOrEmpty(cell?.Text)
					? cell?.Text
					: "")?
					.Length)?
				.Max() ?? 0;

			using (var f = new frmTrimParams())
			{
				f.numMaxLength.Maximum = maxLen;
				f.numMaxLength.Maximum = maxLen;

				if (f.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

				range
					.Cells?
					.Cast<Range>()?
					.ToList()?
					.ForEach(cell => LimitCellTextLength(
						cell,
						(int)f.numMaxLength.Value,
						f.chkTrimStartEnd.Checked,
						f.chkTrimFullSpaces.Checked));
			}
		}

		/// <summary>Trim text to the specifed length</summary>
		/// <param name="value">Text to trim</param>
		/// <param name="len">length to which trim text</param>
		/// <param name="trimStartEndSpaces">Trims start and end spaces of the string</param>
		/// <param name="trimFullSpacedStrings">If source is full of spaces, iy will be trimmed to empty</param>
		private static void LimitCellTextLength(
			Range? cell,
			int len,
			bool trimStartEndSpaces,
			bool trimFullSpacedStrings)
		{
			if (null == cell) return;
			string? value = cell.Text;
			if (string.IsNullOrEmpty(value)) return;

			if (trimFullSpacedStrings && string.IsNullOrWhiteSpace(value))
			{
				cell.Value = string.Empty;
				return;
			};

			value = (value.Length <= len)
				? value
				: value.Substring(0, len);

			if (trimStartEndSpaces) value = value.Trim();

			cell.Value = value;
		}
	}
}
