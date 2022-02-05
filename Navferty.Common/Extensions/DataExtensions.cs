using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class DataExtensions
	{
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataRow> AsEnumerable(this DataRowCollection drc)
			=> drc.Cast<DataRow>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataRow> RowsAsEnumerable(this DataTable dt)
			=> dt.Rows.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataColumn> AsEnumerable(this DataColumnCollection cols)
			=> cols.Cast<DataColumn>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataColumn> ColumnsAsEnumerable(this DataTable dt)
			=> dt.Columns.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewColumn> AsEnumerable(this DataGridViewColumnCollection cols)
			=> cols.Cast<DataGridViewColumn>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewColumn> ColumnsAsEnumerable(this DataGridView grd)
			=> grd.Columns.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewRow> AsEnumerable(this DataGridViewRowCollection rows)
			=> rows.Cast<DataGridViewRow>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewRow> RowsAsEnumerable(this DataGridView grd)
			=> grd.Rows.AsEnumerable();


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewRow> AsEnumerable(this DataGridViewSelectedRowCollection selrows)
			=> selrows.Cast<DataGridViewRow>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewRow> SelectedRowsAsEnumerable(this DataGridView grd)
			=> grd.SelectedRows.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewCell> AsEnumerable(this DataGridViewCellCollection cells)
			=> cells.Cast<DataGridViewCell>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<DataGridViewCell> CellsAsEnumerable(this DataGridViewRow row)
			=> row.Cells.AsEnumerable();
	}
}
