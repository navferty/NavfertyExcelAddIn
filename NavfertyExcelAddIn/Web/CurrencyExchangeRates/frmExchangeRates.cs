using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
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

		private Providers.ExchangeRatesDataProviderBaase ratesProvider = null;

		private System.Data.DataTable dtResult = null;
		private static int ratesDecimalDigitsCount = 2;
		private int columnIndex_Rate = -2;

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
			//Text = string.Format(UIStrings.CurrencyExchangeRates_FormTitle, UIStrings.CurrencyExchangeRates_Sources_CBRF,  ciResult.NumberFormat.CurrencySymbol);
			Text = UIStrings.CurrencyExchangeRates_FormTitle;
			lblSource.Text = UIStrings.CurrencyExchangeRates_Source;
			btnPasteResult.Text = UIStrings.CurrencyExchangeRates_PasteToCell;
			lblFilterTitle.Text = UIStrings.CurrencyExchangeRates_FilterTitle;

			var availProviders = new Providers.ExchangeRatesDataProviderBaase[] {
				new Providers.CBRFProvider(),
				new Providers.NBUProvider()};

			availProviders = (from p in availProviders
							  orderby p.Priority, p.Title
							  select p).ToArray();

			ratesProvider = availProviders.First();
			cbProvider.DataSource = availProviders;
			cbProvider.SelectedIndex = 0;

			var dtNow = DateTime.Now;
			dtpDate.Value = dtNow;
			dtpDate.MaxDate = dtNow;
			dtpDate.MinDate = dtNow.AddYears(-10);
		}

		private async void Form_Displayed(object sender, EventArgs e)
		{
			cbProvider.SelectedIndexChanged += async (s, e) => await UpdateExchangeRates();
			dtpDate.ValueChanged += async (s, e) => await UpdateExchangeRates();
			gridResult.CellFormatting += FormatCell_Rates;

			await UpdateExchangeRates();
			txtFilter.AttachDelayedFilter(() => FilterResultInView(), VistaCueBanner: UIStrings.CurrencyExchangeRates_FilterDescription);
		}

		private async Task UpdateExchangeRates()
		{
			if (cbProvider.SelectedIndex < 0) return;

			this.UseWaitCursor = true;
			try
			{
				ratesProvider = cbProvider.SelectedItem as Providers.ExchangeRatesDataProviderBaase;
				if (ratesProvider == null) return;

				dtResult = null;
				var dtDate = dtpDate.Value;
				{
					var exchangeRatesRows = await ratesProvider.GetExchabgeRatesForDate(dtDate);

					//Count max decimal digits length from all rows
					ratesDecimalDigitsCount = WebResultRow.GetMaxDecimalDigitsCount(exchangeRatesRows);

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
						columnIndex_Rate = gridColumns.ToList().IndexOf(colExchangeRate);
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
				//gridResult.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				dtResult = null;
				gridResult.DataSource = null;
				creator.dialogService.ShowError(ex.Message);
			}
			finally
			{
				FilterResultInView();

				this.UseWaitCursor = false;
				this.Refresh();
			}
		}

		private void FilterResultInView()
		{
			var sFilter = txtFilter.Text.Trim();
			try
			{
				if (dtResult == null) return;

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

			if (e.ColumnIndex != columnIndex_Rate) return;
			if (e.Value is not double dRate) return;

			e.Value = dRate.ToString($"C{ratesDecimalDigitsCount}", ratesProvider.Culture);
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
					creator.dialogService.ShowError(UIStrings.CurrencyExchangeRates_Error_NedAnyCellSelection);
					return;
				}

				var selRows = gridResult.SelectedRowsAsEnumerable();
				if (selRows.Count() != 1)
				{
					creator.dialogService.ShowError(UIStrings.CurrencyExchangeRates_Error_CanSelectOnlyOneRow);
					return;
				}

				var selRow = selRows.First();
				var exchangeRate = Convert.ToDecimal(selRow.CellsAsEnumerable().ToArray()[columnIndex_Rate].Value);
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
