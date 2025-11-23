using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.Undo
{
	/// <summary>
	/// Allows to undo actions invoked by
	/// <see cref="EnumerableExtensions.ApplyForEachCellOfType"/>
	/// </summary>
	public class UndoManager
	{
		private readonly IList<UndoItem> undoItems = new List<UndoItem>();
		private readonly IList<AreaUndoItems> areaUndoItems = new List<AreaUndoItems>();

		private string? wsName;
		private string? wbName;

		public void StartNewAction(Range range)
		{
			undoItems.Clear();
			areaUndoItems.Clear();
			wsName = range.Worksheet.Name;
			wbName = (string)range.Worksheet.Parent.Name;
		}

		public void PushUndoItem(UndoItem undoItem)
		{
			undoItems.Add(undoItem);
		}

		public void PushAreaUndoItem(AreaUndoItems undoItem)
		{
			areaUndoItems.Add(undoItem);
		}

		public void UndoLastAction(Worksheet worksheet)
		{
			if (!undoItems.Any() && !areaUndoItems.Any())
				return;

			if (wsName != worksheet.Name)
				return;

			if (wbName != (string)worksheet.Parent.Name)
				return;

			if (undoItems.Any(x => CheckChanges(worksheet, x)))
				return;

			if (areaUndoItems
					.Any(x => CheckChanges(worksheet, x)))
				return;

			foreach (var item in undoItems)
			{
				worksheet.Cells[item.RowIndex, item.ColumnIndex].Value = item.OldValue;
			}

			foreach (var item in areaUndoItems)
			{
				var range = GetRange(worksheet, item);
				range.Value = item.OldValues;
			}

			undoItems.Clear();
		}

		private static Range GetRange(Worksheet worksheet, AreaUndoItems item)
		{
			return worksheet.Range[
				worksheet.Cells[item.Row, item.Column],
				worksheet.Cells[item.Row + item.Height - 1, item.Column + item.Width - 1]];
		}

		private static bool CheckChanges(Worksheet worksheet, UndoItem x)
		{
			object? currentValue = ((Range)worksheet.Cells[x.RowIndex, x.ColumnIndex]).Value;

			if (currentValue?.GetType() != x.NewValue?.GetType())
				return true;

			if (x.NewValue is null)
				return currentValue is null;

			return x.NewValue.Equals(currentValue);
		}

		private static bool CheckChanges(Worksheet worksheet, AreaUndoItems areaUndoItems)
		{
			var range = GetRange(worksheet, areaUndoItems);
			var values = areaUndoItems.NewValues;

			if (!(range.Value is object[,] currentValue))
				return true;

			int upperI = currentValue.GetUpperBound(0); // Rows
			int upperJ = currentValue.GetUpperBound(1); // Columns

			if (upperI != values.GetUpperBound(0)
				|| upperJ != values.GetUpperBound(1)
				|| currentValue.GetLowerBound(0) != values.GetLowerBound(0)
				|| currentValue.GetLowerBound(1) != values.GetLowerBound(1))
				return true;

			for (int i = currentValue.GetLowerBound(0); i <= upperI; i++)
			{
				for (int j = currentValue.GetLowerBound(1); j <= upperJ; j++)
				{
					if (currentValue[i, j] == null && values[i, j] == null)
						continue;

					if (currentValue[i, j] == null || values[i, j] == null)
						return true;

					if (!currentValue[i, j].Equals(values[i, j]))
						return true;
				}
			}

			return false;
		}
	}

	[SuppressMessage("Performance", "CA1819:Properties should not return arrays",
		Justification = "Excel return range values as double-dimensioned array")]
	public class AreaUndoItems
	{
		public int Row { get; set; }
		public int Column { get; set; }
		public int Width { get; set; }
		public int Height { get; set; }

		public object?[,] OldValues { get; set; } = null!;
		public object?[,] NewValues { get; set; } = null!;
	}

	public class UndoItem
	{
		public object? OldValue { get; set; }
		public object? NewValue { get; set; }

		public int RowIndex { get; set; }
		public int ColumnIndex { get; set; }
	}
}
