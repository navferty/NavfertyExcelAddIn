using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

using DataTable = System.Data.DataTable;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public partial class frmExchangeRates : Form
	{
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		private readonly CurrencyExchangeRates creator = null;
		private readonly Workbook wb = null;


		private static readonly CultureInfo ciRU = CultureInfo.GetCultureInfo("ru-RU");
		private static readonly CultureInfo ciUA = CultureInfo.GetCultureInfo("uk-UA");
		//private static readonly CultureInfo ciUS = CultureInfo.GetCultureInfo("en-US");

		private static readonly Dictionary<string, uint> vipCurrencies = new()
		{
			{ "USD", 1u },
			{ "EUR", 2u },
			{ "GBP", 3u },
			{ "CNY", 4u }
		};
		private static readonly DataGridViewCellStyle cellStyle_ExchangeRate = new() { Alignment = DataGridViewContentAlignment.MiddleRight };

		private static string columnCurrencyTitle = "Валюта";
		private static string columnCurrencyCode = "Код";
		private static string columnCurrencyRate = "Курс";

		private System.Data.DataTable dtResult = null;
		private static int exchangeRatesDecimalDigitsCount = 2;
		private CultureInfo ciResult = ciRU;
		private int columnIndex_ExchangeRate = -2;

		public frmExchangeRates()
		{
			InitializeComponent();
		}
		public frmExchangeRates(CurrencyExchangeRates Creator, Workbook wb) : this()
		{
			this.creator = Creator;
			this.wb = wb;
		}

		private void Form_Load(object sender, EventArgs e)
		{
			Text = string.Format(UIStrings.CurrencyExchangeRates_FormTitle, UIStrings.CurrencyExchangeRates_Sources_CBRF, ciResult.NumberFormat.CurrencySymbol);
			btnPasteResult.Text = UIStrings.CurrencyExchangeRates_PasteToCell;

			var dtNow = DateTime.Now;
			dtpDate.Value = dtNow;
			dtpDate.MaxDate = dtNow;
			dtpDate.MinDate = dtNow.AddYears(-10);
		}

		private async void Form_Displayed(object sender, EventArgs e)
		{
			dtpDate.ValueChanged += DtpDate_ValueChanged;
			gridResult.CellFormatting += FormatCell_Rates;

			await UpdateExchangeRates();
			txtFilter.AttachDelayedFilter(() => FilterResultInView(), VistaCueBanner: UIStrings.CurrencyExchangeRates_FilterTitle);
		}

		private async void DtpDate_ValueChanged(object sender, EventArgs e)
		{
			await UpdateExchangeRates();
		}

		private async Task UpdateExchangeRates()
		{
			this.UseWaitCursor = true;
			try
			{
				dtResult = null;
				var dtDate = dtpDate.Value;
				{
					var exchangeRatesRows = await CurrencyExchangeRates.GetCurrencyExchabgeRates_CBRF(dtDate);

					//Count max decimal digits length from all rows
					exchangeRatesDecimalDigitsCount = WebResultRow.GetMaxDecimalDigitsCount(exchangeRatesRows);

					//Sort by priority
					exchangeRatesRows.ToList().ForEach(wrr =>
					{
						var bIsVIPCurrency = vipCurrencies.TryGetValue(wrr.ISOCode, out uint iPriorityFound);
						if (bIsVIPCurrency) wrr.PriorityInGrid = iPriorityFound;
					});

					exchangeRatesRows = (from r in exchangeRatesRows
										 orderby r.PriorityInGrid ascending, r.Name ascending
										 select r).ToArray();


					var dtView = new DataTable();
					{
						DataColumn colName = new(columnCurrencyTitle, typeof(string));
						DataColumn colISO3 = new(columnCurrencyCode, typeof(string));
						DataColumn colExchangeRate = new(columnCurrencyRate, typeof(double));
						//TODO: добавить колонку даты актуальности для строк

						var gridColumns = new[] { colName, colISO3, colExchangeRate };
						dtView.Columns.AddRange(gridColumns);
						columnIndex_ExchangeRate = gridColumns.ToList().IndexOf(colExchangeRate);
					}

					foreach (var old in exchangeRatesRows)
					{
						var newRow = dtView.NewRow();
						newRow.ItemArray = new object[] { old.FullNameWithUnits, old.ISOCode, old.Curs };
						dtView.Rows.Add(newRow);
					}

					dtResult = dtView;
				};

				gridResult.DataSource = dtResult;
				if (dtResult == null) return;

				if (gridResult.Columns.Count > 0)
				{
					var colRate = gridResult.ColumnsAsEnumerable().Last();
					if (colRate.DefaultCellStyle != cellStyle_ExchangeRate) colRate.DefaultCellStyle = cellStyle_ExchangeRate;
				}
				FilterResultInView();
				//gridResult.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				gridResult.DataSource = null;
				//Commons.DialogService.ReferenceEquals
				MessageBox.Show(ex.Message);
			}
			finally
			{
				this.UseWaitCursor = false;
				this.Refresh();
			}
		}

		private void FilterResultInView()
		{
			var sFilter = txtFilter.Text.Trim();
			try
			{
				if (string.IsNullOrWhiteSpace(sFilter))
				{
					sFilter = string.Empty;
				}
				else
				{
					var columnNames = dtResult.ColumnsAsEnumerable()
						.Select(col =>
						{
							if (col.DataType == typeof(string)) //We can filter only text fields
								return $"[{col.ColumnName}] LIKE '%{sFilter}%'";

							return "";
						})
						.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

					sFilter = string.Join(" OR ", columnNames);
					//Debug.WriteLine("Row filter = " + sFilter);
				}
				dtResult.DefaultView.RowFilter = sFilter;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Apply Row filter ('{sFilter}') ERROR!\n" + ex.Message);
			}
			finally
			{
				UpdatePasteButtonState();
			}
		}

		/// <summary> Format Rate column cells like number with thouthand separator</summary>
		private void FormatCell_Rates(object sender, DataGridViewCellFormattingEventArgs e)
		{
			if (e.Value == null || e.RowIndex == gridResult.NewRowIndex) return;

			if (e.ColumnIndex != columnIndex_ExchangeRate) return;
			if (e.Value is not double dRate) return;

			e.Value = dRate.ToString($"C{exchangeRatesDecimalDigitsCount}", ciResult);
		}

		private void UpdatePasteButtonState()
		{
			btnPasteResult.Enabled = gridResult.RowsAsEnumerable().Any();
		}

		private void gridResult_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 0) return;
			OnPasteResult();
		}

		private void btnPasteResult_Click(object sender, EventArgs e)
		{
			OnPasteResult();
		}

		private void OnPasteResult()
		{
			try
			{
				if (!btnPasteResult.Enabled) return;

				if (App.Selection == null
					|| ((Range)App.Selection).Cells == null
					|| ((Range)App.Selection).Cells.Count != 1)
				{
					creator.dialogService.ShowError(UIStrings.CurrencyExchangeRates_NedAnyCellSelection);
					return;
				}

				var selRows = gridResult.SelectedRowsAsEnumerable();
				if (selRows.Count() != 1)
				{
					creator.dialogService.ShowError("Надо выбрать только одну строку для вставки!");
					return;
				}

				var selRow = selRows.First();
				var exchangeRate = Convert.ToDecimal(selRow.CellsAsEnumerable().ToArray()[columnIndex_ExchangeRate].Value);
				//TODO: Учитывать количество Units при выдаче курса!

				Range selectedExcelRange = (Range)App.Selection;
				selectedExcelRange.Value = exchangeRate;


				DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
		}
	}
}
