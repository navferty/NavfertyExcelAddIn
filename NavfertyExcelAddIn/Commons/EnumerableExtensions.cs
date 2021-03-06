﻿using System;
using System.Collections.Generic;
using System.Linq;

using Autofac;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Undo;

using NLog;

namespace NavfertyExcelAddIn.Commons
{
	public static class EnumerableExtensions
	{
		private static readonly UndoManager undoManager =
			NavfertyRibbon.Container.Resolve<UndoManager>();

		private static readonly ILogger logger = LogManager.GetCurrentClassLogger();

		public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
		{
			foreach (T element in source)
			{
				action(element);
			}
		}

		public static void ForEachCell(this Range range, Action<Range> action)
		{
			// TODO rewrite to use less read-write calls to interop (like Range.Value)
			range.Cast<Range>().ForEach(action);
		}

		public static void ApplyForEachCellOfType<TIn, TOut>(this Range range, Func<TIn, TOut> transform)
		{
			logger.Debug($"Apply transformation to range '{range.GetRelativeAddress()}' on worksheet '{range.Worksheet.Name}'");

			undoManager.StartNewAction(range);

			foreach (Range area in range.Areas)
			{
				ApplyToArea(area, transform);
			}
		}

		private static void ApplyToArea<TIn, TOut>(Range range, Func<TIn, TOut> transform)
		{
			var rangeValue = range.Value;
			if (rangeValue is null)
				return;

			if (rangeValue is TIn currentValue)
			{
				var newValue = transform(currentValue);
				range.Value = newValue;
				var undoItem = new UndoItem
				{
					OldValue = currentValue,
					NewValue = newValue,
					ColumnIndex = range.Column,
					RowIndex = range.Row
				};
				undoManager.PushUndoItem(undoItem);
				return;
			}

			// minimize number of COM calls to excel
			if (!(rangeValue is object[,] values))
				return;

			int upperI = values.GetUpperBound(0); // Rows
			int upperJ = values.GetUpperBound(1); // Columns

			var isChanged = false;
			var oldValues = (object[,])values.Clone();

			logger.Debug($"Converting columns from {values.GetLowerBound(0)} to {upperI}, " +
				$"rows from {values.GetLowerBound(1)} to {upperJ}");

			for (int i = values.GetLowerBound(0); i <= upperI; i++)
			{
				for (int j = values.GetLowerBound(1); j <= upperJ; j++)
				{
					var value = values[i, j];
					if (value is TIn s)
					{
						var newValue = transform(s);
						if ((object)newValue != value) // TODO check boxing time on million values
						{
							isChanged = true;
							values[i, j] = newValue;
						}
					}
				}
			}

			if (isChanged)
			{
				logger.Debug("Some values were converted, writing to worksheet");
				range.Value = values;
				undoManager.PushAreaUndoItem(new AreaUndoItems
				{
					NewValues = values,
					OldValues = oldValues,
					Column = range.Column,
					Row = range.Row,
					Height = upperI,
					Width = upperJ
				});
			}
		}
	}
}
