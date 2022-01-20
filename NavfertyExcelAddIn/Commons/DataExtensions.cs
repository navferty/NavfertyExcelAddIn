using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Windows.Forms;
using System.Runtime.CompilerServices;

namespace NavfertyExcelAddIn.Commons
{
	internal static class DataExtensions
	{

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Data.DataRow> AsEnumerable(this System.Data.DataRowCollection drc)
			=> drc.Cast<System.Data.DataRow>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Data.DataRow> RowsAsEnumerable(this System.Data.DataTable dt)
			=> dt.Rows.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Data.DataColumn> AsEnumerable(this System.Data.DataColumnCollection cols)
			=> cols.Cast<System.Data.DataColumn>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Data.DataColumn> ColumnsAsEnumerable(this System.Data.DataTable dt)
			=> dt.Columns.AsEnumerable();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Windows.Forms.DataGridViewColumn> AsEnumerable(this System.Windows.Forms.DataGridViewColumnCollection cols)
			=> cols.Cast<System.Windows.Forms.DataGridViewColumn>();

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static IEnumerable<System.Windows.Forms.DataGridViewColumn> ColumnsAsEnumerable(this System.Windows.Forms.DataGridView grd)
			=> grd.Columns.AsEnumerable();




	}
}
