using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using Navferty.Common;
using Navferty.ExcelAddIn.Web.Localization;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates
{
	internal partial class frmExchangeRates : Navferty.Common.Controls.FormEx
	{
		private readonly IDialogService? dialogService;

		private static readonly Dictionary<string, uint> vipCurrencies = new()
		{
			{ "USD", 1u },
			{ "EUR", 2u },
			{ "GBP", 3u },
			{ "CNY", 4u }
		};

		private readonly DataGridViewCellStyle cellStyle_ExchangeRate = new() { Alignment = DataGridViewContentAlignment.MiddleRight };

		private const string GRID_COLUMNS_NAME = "Name";
		private const string GRID_COLUMNS_ISO = "ISO";
		private const string GRID_COLUMNS_RATE = "Rate";
		private readonly Lazy<Dictionary<string, string>> dicGridColumnTitlesLazy = new(() => new Dictionary<string, string>() {
			{ GRID_COLUMNS_NAME, UIStrings.CurrencyExchangeRates_GridColumn_Name },
			{ GRID_COLUMNS_ISO, UIStrings.CurrencyExchangeRates_GridColumn_ISO },
			{ GRID_COLUMNS_RATE, UIStrings.CurrencyExchangeRates_GridColumn_Rate }
		});

		private Providers.ExchangeRatesProviderBase? ratesProvider;
		private CurrencyExchangeRatesDataset.ExchangeRatesDataTable? dtResult;
		private static int ratesDecimalDigitsCount = 4;


		public ExchangeRateRecord? SelectedExchangeRate;

		/// <summary>Just for designer</summary>
		[EditorBrowsable(EditorBrowsableState.Never)]
		public frmExchangeRates()
		{
			InitializeComponent();
		}

		public frmExchangeRates(IDialogService ds) : this()
		{
			dialogService = ds;
		}

		private void Form_Load(object sender, EventArgs e)
		{
			Text = UIStrings.CurrencyExchangeRates_FormTitle;
			lblInfoTip.Text = UIStrings.CurrencyExchangeRates_FormInfoTip;
			lblInfoTip.Visible = false;


			lblSource.Text = UIStrings.CurrencyExchangeRates_Source;
			btnPasteResult.Text = UIStrings.CurrencyExchangeRates_PasteToCell;
			lblFilterTitle.Text = UIStrings.CurrencyExchangeRates_FilterTitle;

			var availProviders = new Providers.ExchangeRatesProviderBase[] {
				new Providers.CBRF.CBRFProvider(),
				new Providers.NBU.NBUProvider(),
				new Providers.ECB.ECBProvider()};

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

			gridResult.EmptyText = "Загружаем данные из сети...";
			gridResult.Refresh();

			this.UseWaitCursor = true;
			try
			{
				ratesProvider = cbProvider.SelectedItem as Providers.ExchangeRatesProviderBase;
				if (ratesProvider == null) return;

				dtResult = null;
				var dtDate = dtpDate.Value.Date;
				{
					dtResult = await ratesProvider.GetExchangeRatesForDate(dtDate, wrr =>
					{
						var bIsVIPCurrency = vipCurrencies.TryGetValue(wrr.ISOCode, out uint iPriorityFound);
						return bIsVIPCurrency
						? iPriorityFound
						: null;
					});


					// Count max decimal digits length from all rows
					//ratesDecimalDigitsCount = WebResultRow.GetMaxDecimalDigitsCount(exchangeRatesRows);

					//Some rows (in Ukrainian NBU) has rates like 0.00000000698  and than all rows in grid looks weird.
					//To avoid this, we use standart 4 digits float part
				};

				gridResult.DataSource = dtResult;
				if (dtResult == null) return;

				if (gridResult.Columns.Count > 0)
				{
					gridResult.ColumnsAsEnumerable().ToList().ForEach(col =>
					{
						var bfound = dicGridColumnTitlesLazy.Value.TryGetValue(col.Name, out string FoundTitle);
						col.Visible = bfound;//Hide columns that have not translated titles (this is raw helpers data)
						if (bfound) col.HeaderText = FoundTitle;


						if (col.Name == GRID_COLUMNS_RATE)
						{
							if (col.DefaultCellStyle != cellStyle_ExchangeRate) col.DefaultCellStyle = cellStyle_ExchangeRate;
						}
					});
				}
			}
			catch (Exception ex)
			{
				dtResult = null;
				gridResult.EmptyText = ex.Message;
				gridResult.DataSource = null;

				ratesProvider?.Logger.Error(ex);
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

				gridResult.EmptyText = "Данные успешно загружены. Нет строк для отображения";

				if (string.IsNullOrWhiteSpace(sFilter))
				{
					sFilter = string.Empty;
				}
				else
				{
					gridResult.EmptyText = "Нет строк соответствующих указанному фильтру.";

					var columnNames = dtResult.ColumnsAsEnumerable()
						.Select(col =>
						{
							var bIsColumnVisible = dicGridColumnTitlesLazy.Value.TryGetValue(col.ColumnName, out string FoundTitle);
							if (bIsColumnVisible)//Filter only visible columns
							{
								if (col.DataType == typeof(string)) //We can filter only text fields
									return $"[{col.ColumnName}] LIKE '%{sFilter}%'";

							}
							return "";
						})
						.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

					sFilter = string.Join(" OR ", columnNames);
					Debug.WriteLine($"Row filter: {sFilter}");
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

		private int GetGridColumnIdex(string colID)
		{
			var columnIDs = gridResult.ColumnsAsEnumerable().Select(c => c.Name).ToArray();
			return Array.IndexOf(columnIDs, colID);
		}

		/// <summary> Format Rate column cells like number with thouthand separator</summary>
		private void FormatCell_Rates(object sender, DataGridViewCellFormattingEventArgs e)
		{
			if (e.Value == null || e.RowIndex == gridResult.NewRowIndex) return;

			var iColRate = GetGridColumnIdex(GRID_COLUMNS_RATE);
			if (e.ColumnIndex != iColRate) return;
			if (e.Value is not double dRate) return;

			e.Value = dRate.ToString($"C{ratesDecimalDigitsCount}", ratesProvider!.Culture);
		}

		private void UpdatePasteButtonState()
		{
			btnPasteResult.Enabled = gridResult.RowsAsEnumerable().Any();
			lblInfoTip.Visible = gridResult.RowsAsEnumerable().Any();
		}

		private void gridResult_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 0) return;
			OnPasteResultToWorkSheet();
		}

		private void btnPasteResult_Click(object sender, EventArgs e)
		{
			OnPasteResultToWorkSheet();
		}

		private void OnPasteResultToWorkSheet()
		{
			try
			{
				if (!btnPasteResult.Enabled) return;


				var selRows = gridResult.SelectedRowsAsEnumerable();
				if (selRows.Count() != 1)
				{
					dialogService?.ShowError(UIStrings.CurrencyExchangeRates_Error_CanSelectOnlyOneRow);
					return;
				}

				var selRow = selRows.First();
				var err = ((selRow.DataBoundItem as DataRowView)!.Row as CurrencyExchangeRatesDataset.ExchangeRatesRow);
				this.SelectedExchangeRate = err!.Raw as ExchangeRateRecord;
				DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				dialogService?.ShowError(ex.Message);
			}
		}


	}
}
