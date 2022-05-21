using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

using Microsoft.Office.Interop.Excel;

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class RangeExtensions
	{
		public const Int64 DEFAULT_MAX_ALLOWED_CELLS = 2;// 1_000_000L;

		public class TooManyCellsException : Exception
		{
			public TooManyCellsException() : base(Localization.UIStrings.Error_TooManyCellsSelected) { }
		}

		/// <summary>Throws error if cells count in specifed range more than 'maxAllowedCellsInRange' (DEFAULT_MAX_ALLOWED_CELLS_IN_RANGE)
		/// </summary>
		/// <param name="maxAllowedCellsInRange"></param>
		/// <exception cref="TooManyCellsException"></exception>

		[DebuggerStepThrough]
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static Int64 ThrowIfTooManyCellsSelected(
			this Range? sel,
			Int64 maxAllowedCellsInRange = DEFAULT_MAX_ALLOWED_CELLS)
		{
			Int64 iCellsSelected = sel?.Cells?.Count ?? 0;

			if (iCellsSelected > DEFAULT_MAX_ALLOWED_CELLS)
				throw new TooManyCellsException();

			return iCellsSelected;
		}
	}
}
