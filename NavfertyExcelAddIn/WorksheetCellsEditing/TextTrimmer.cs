using System;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

using NLog;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public class TextTrimmer : ITextTrimmer
	{
		private static readonly ILogger logger = LogManager.GetCurrentClassLogger();

		public void TrimExtraSpaces(Range range)
		{
			logger.Info($"Trim spaces for range {range.GetRelativeAddress()}");

			range.ApplyForEachCellOfType<string, string>(value => value.TrimSpaces());
		}

		public void RemoveAllSpaces(Range range)
		{
			logger.Info($"Trim spaces for range {range.GetRelativeAddress()}");

			range.ApplyForEachCellOfType<string, string>(value =>
			{
				if (string.IsNullOrWhiteSpace(value))
					return null;

				return new string(value.Where(c => !char.IsWhiteSpace(c)).ToArray());
			});
		}

		public void TrimTextByLengthUIDisplay(Range range)
		{

			range.ThrowIfTooManyCellsSelected();

			int maxLen = range
				.Cells?
				.Cast<Range>()?
				.Select(cell => (string.IsNullOrEmpty(cell?.Text)
					? "" : cell?.Text
					)?
				.Length)?.Max() ?? 0;

			using (var f = new TrimTextByLengthUI.frmTrimParams())
			{
				f.numMaxLength.Maximum = maxLen;
				f.numMaxLength.Value = maxLen;
				f.numMaxLength.Enabled = (maxLen > 0);


				f.Text = UIStrings.TrimTextByLength_Title;
				f.lblTextLength.Text = UIStrings.TrimTextByLength_TextLen;
				f.chkTrimStartEnd.Text = UIStrings.TrimTextByLength_TrimStartEnd;
				f.chkTrimFullSpaces.Text = UIStrings.TrimTextByLength_TrimFullSpaced;

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
