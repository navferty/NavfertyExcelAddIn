using System;

using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public class ConditionalFormatFixer : IConditionalFormatFixer
	{
		public void FillRange(Range range)
		{
			var rowsCount = range.Rows.Count;
			if (rowsCount < 2)
			{
				// logger
				return;
			}

			Range firstRow = range.Rows[1];
			Range secondRow = range.Rows[2];

			if (rowsCount == 2)
			{
				if (firstRow.FormatConditions.Count == 0
					&& secondRow.FormatConditions.Count != 0)
				{
					CopyConditionalFormat( source: secondRow, target: firstRow);
				}
				else
				{
					secondRow.FormatConditions.Delete();
					CopyConditionalFormat( source: firstRow, target: secondRow);
				}

				return;
			}

			if (firstRow.FormatConditions.Count == 0)
			{
				firstRow = secondRow;
				secondRow = range.Rows[3];
			}

			Range lastRow = range.Rows[rowsCount];

			var ws = range.Worksheet;
			ws.Range[secondRow, lastRow].FormatConditions.Delete();

			CopyConditionalFormat(
				
				source: firstRow,
				target: ws.Range[firstRow, lastRow]);
		}

		private static void CopyConditionalFormat(Range source, Range target)
		{
			source.Copy();
			target.PasteSpecial(XlPasteType.xlPasteFormats);
			source.Worksheet.Application.CutCopyMode = 0;
		}
	}
}
