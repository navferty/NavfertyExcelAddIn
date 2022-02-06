using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Autofac;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.DataValidation;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.InteractiveRangeReport;
using NavfertyExcelAddIn.Localization;
using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.StringifyNumerics;
using NavfertyExcelAddIn.Transliterate;
using NavfertyExcelAddIn.Undo;
using NavfertyExcelAddIn.UnprotectWorkbook;
using NavfertyExcelAddIn.WorksheetCellsEditing;
using NavfertyExcelAddIn.WorksheetProtectUnprotect;
using NavfertyExcelAddIn.XmlTools;

using NLog;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn
{
	[ComVisible(true)]
	[SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Ribbon callbacks must have certain signature")]
	[SuppressMessage("Design", "CA1063:Implement IDisposable Correctly", Justification = "Class contains only manages disposable resources")]
	public class NavfertyRibbon : IRibbonExtensibility, IDisposable
	{
		#region Public static members
		public static readonly IContainer Container = Registry.CreateContainer();
		#endregion

		#region Private members
		private readonly ILogger logger = LogManager.GetCurrentClassLogger();
		private readonly IDialogService dialogService = Container.Resolve<IDialogService>();

		private readonly Dictionary<string, ValidationType> validationTypeByButtonId =
			new Dictionary<string, ValidationType>
			{
				{ "ValidateValuesNumerics", ValidationType.Numeric },
				{ "ValidateValuesXml", ValidationType.Xml },
				{ "ValidateValuesDate", ValidationType.Date },
				{ "ValidateValuesTinPersonal", ValidationType.TinPersonal },
				{ "ValidateValuesTinOrganization", ValidationType.TinOrganization }
			};

		private Application App => Globals.ThisAddIn.Application;

		#region Forms
		private InteractiveRangeReportForm form;
		#endregion
		#endregion

		#region IRibbonExtensibility
		public string GetCustomUI(string ribbonID)
		{
			var sw = Stopwatch.StartNew();
			logger.Debug($"{nameof(GetCustomUI)} - {ribbonID}");

			var customUiXml = GetResourceText();

			logger.Debug($"{nameof(GetCustomUI)} got xml for {sw.Elapsed}");
			return customUiXml;
		}
		#endregion

		#region Ribbon callbacks
		public void RibbonLoad(IRibbonUI ribbonUI)
		{
			logger.Debug($"Ribbon loaded");
		}

		#region View Group
		public void ToggleSheetLabels(IRibbonControl ribbonControl)
		{
			if (App.ActiveWindow != null)
				App.ActiveWindow.DisplayWorkbookTabs = !App.ActiveWindow.DisplayWorkbookTabs;
		}
		public void SwitchReferenceStyle(IRibbonControl ribbonControl)
		{
			if (App.ActiveWindow == null)
				return;

			switch (App.ReferenceStyle)
			{
				case XlReferenceStyle.xlA1:
					App.ReferenceStyle = XlReferenceStyle.xlR1C1;
					break;
				case XlReferenceStyle.xlR1C1:
					App.ReferenceStyle = XlReferenceStyle.xlA1;
					break;
			}
		}
		#endregion

		#region Common tools
		public void UndoLastAction(IRibbonControl ribbonControl)
		{
			var undoManager = GetService<UndoManager>();
			undoManager.UndoLastAction(App.ActiveSheet);
		}

		public void ParseNumerics(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"Parse numerics for range {range.Address}");

			var parser = GetService<INumericParser>();
			parser.Parse(range);
		}

		public void NumberToWordsEnglish(IRibbonControl ribbonControl) =>
			StringifyNumerics(SupportedCulture.English);
		public void NumberToWordsRussian(IRibbonControl ribbonControl) =>
			StringifyNumerics(SupportedCulture.Russian);
		public void NumberToWordsFrench(IRibbonControl ribbonControl) =>
			StringifyNumerics(SupportedCulture.French);

		private void StringifyNumerics(SupportedCulture supportedCulture)
		{
			var stringifier = Container.ResolveKeyed<INumericStringifier>(supportedCulture);

			Range selection = GetSelectionOrUsedRange(App.ActiveSheet);
			selection.ApplyForEachCellOfType<double, object>(
				value =>
				{
					var newValue = stringifier.StringifyNumber(value);
					return (object)newValue ?? value;
				});
		}

		public void ReplaceChars(IRibbonControl ribbonControl)
		{
			var replacer = GetService<ICyrillicLettersReplacer>();

			Range selection = GetSelectionOrUsedRange(App.ActiveSheet);
			selection.ApplyForEachCellOfType<string, object>(
				value =>
				{
					var newValue = replacer.ReplaceCyrillicCharsWithLatin(value);
					return (object)newValue ?? value;
				});
		}

		public void Transliterate(IRibbonControl ribbonControl)
		{
			var transliterator = GetService<ITransliterator>();

			Range selection = GetSelectionOrUsedRange(App.ActiveSheet);
			selection.ApplyForEachCellOfType<string, object>(
				value =>
				{
					var newValue = transliterator.Transliterate(value);
					return (object)newValue ?? value;
				});
		}

		public void UnprotectWorkbook(IRibbonControl ribbonControl)
		{
			var wb = App.ActiveWorkbook;
			var path = wb.FullName;

			logger.Debug($"UnprotectWorkbook {path}");

			var extension = path.Split('.').LastOrDefault();

			if (extension != "xlsx" && extension != "xlsm")
			{
				dialogService.ShowError(UIStrings.CannotUnlockPleaseSaveAsXml);
				return;
			}

			if (!dialogService.Ask(UIStrings.UnsavedChangesWillBeLostPrompt, UIStrings.Warning))
			{
				return;
			}

			wb.Close(false);


			var wbUnprotector = GetService<IWbUnprotector>();
			wbUnprotector.UnprotectWorkbookWithAllWorksheets(path);

			App.Workbooks.Open(path);
		}

		public void ProtectUnprotectWorksheets(IRibbonControl ribbonControl)
		{
			var wb = App.ActiveWorkbook;
			if (wb is null) return;

			var path = wb.FullName;
			logger.Debug($"ProtectUnprotectWorksheets {path}");

			var wsProtectorUnprotector = GetService<IWsProtectorUnprotector>();
			wsProtectorUnprotector.ProtectUnprotectSelectedWorksheets(wb);
		}

		public void CutNames(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"CutNames. Range selected is {range.Address}");

			// TODO
		}

		public void HighlightDuplicates(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"HighlightDuplicates. Range selected is {range.Address}");


			var duplicatesHighlighter = GetService<IDuplicatesHighlighter>();
			duplicatesHighlighter.HighlightDuplicates(range);
		}

		public void ToggleCase(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"ToggleCase. Range selected is {range.Address}");


			var caseToggler = GetService<ICaseToggler>();
			caseToggler.ToggleCase(range);
		}

		public void TrimSpaces(IRibbonControl ribbonControl)
		{
			TrimExtraSpaces(ribbonControl);
		}

		public void TrimExtraSpaces(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"{nameof(TrimExtraSpaces)}. Range selected is {range.Address}");


			var trimmer = GetService<IEmptySpaceTrimmer>();
			trimmer.TrimExtraSpaces(range);
		}

		public void RemoveAllSpaces(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"{nameof(RemoveAllSpaces)}. Range selected is {range.Address}");


			var trimmer = GetService<IEmptySpaceTrimmer>();
			trimmer.RemoveAllSpaces(range);
		}

		public void RepairConditionalFormat(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);
			if (range == null)
				return;

			logger.Debug($"{nameof(RepairConditionalFormat)}. Range selected is {range.Address}");

			var formatFixer = GetService<IConditionalFormatFixer>();
			formatFixer.FillRange(range);
		}


		public void UnmergeCells(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"UnmergeCells. Range selected is {range.Address}");


			var cellsUnmerger = GetService<ICellsUnmerger>();
			cellsUnmerger.Unmerge(range);

		}

		public void ValidateValues(IRibbonControl ribbonControl)
		{
			var activeSheet = (Worksheet)App.ActiveSheet;
			var range = GetSelectionOrUsedRange(activeSheet);

			if (range == null)
				return;

			logger.Debug($"ValidateValues. Range selected is {range.Address}");

			if (!validationTypeByButtonId.TryGetValue(ribbonControl.Id, out var validationType))
			{
				dialogService.ShowError($"Invalid control id '{ribbonControl.Id}'");
				throw new ArgumentOutOfRangeException($"Invalid control id '{ribbonControl.Id}'");
			}

			logger.Debug($"ValidateValues. Range selected is {range.Address}, validation type {validationType}");

			IReadOnlyCollection<InteractiveErrorItem> results;

			var validator = GetService<ICellsValueValidator>();
			results = validator.Validate(range, validationType);

			form = new InteractiveRangeReportForm(results, activeSheet);
			form.Show();
		}

		public void FindErrors(IRibbonControl ribbonControl)
		{
			var activeSheet = (Worksheet)App.ActiveSheet;
			var range = GetSelectionOrUsedRange(activeSheet);

			if (range == null)
				return;

			if (range.Count == 1)
				range = activeSheet.UsedRange;

			logger.Debug($"FindErrors. Range selected is {range.Address}");

			IReadOnlyCollection<InteractiveErrorItem> allErrors;

			var errorFinder = GetService<IErrorFinder>();
			allErrors = errorFinder.GetAllErrorCells(range);

			if (allErrors.Count == 0)
			{
				dialogService.ShowInfo(string.Format(UIStrings.NoErrors, range.Address));
				return;
			}
			form = new InteractiveRangeReportForm(allErrors, activeSheet);
			form.Show();
		}

		public void CopyAsMarkdown(IRibbonControl ribbonControl)
		{
			var range = GetSelectionOrUsedRange(App.ActiveSheet);

			if (range == null)
				return;

			logger.Debug($"CopyAsMarkdown. Range selected is {range.Address}");

			string table;

			var markdownReader = GetService<ICellsToMarkdownReader>();
			table = markdownReader.ReadToMarkdown(range);

			if (!string.IsNullOrWhiteSpace(table))
			{
				Clipboard.SetText(table);
			}
		}
		#endregion

		#region XML Tools
		public void CreateSampleXml(IRibbonControl ribbonControl)
		{
			logger.Debug("CreateSampleXml pressed");


			var xmlSampleCreator = GetService<IXmlSampleCreator>();
			xmlSampleCreator.CreateSampleXml();
		}
		public void ValidateXml(IRibbonControl ribbonControl)
		{
			logger.Debug("ValidateXml pressed");


			var validator = GetService<IXmlValidator>();
			validator.Validate(App);
		}
		#endregion

		#region Web
		public void CurrencyExchangeRatesSelect(IRibbonControl ribbonControl)
		{
			var wb = App.ActiveWorkbook;
			if (wb is null) return;

			var path = wb.FullName;
			logger.Debug($"Web.CurrencyExchangeRates {path}");

			var webExchangeRates = GetService<Web.IWebTools>();
			webExchangeRates.CurrencyExchangeRates_Show();
		}
		#endregion

		#endregion

		#region Utils
		public string GetLabel(IRibbonControl ribbonControl)
		{
			var sw = Stopwatch.StartNew();
			logger.Debug($"{nameof(GetLabel)} - {ribbonControl?.Id}");
			try
			{
				var label = RibbonLabels.ResourceManager.GetString(ribbonControl.Id);
				logger.Debug($"{nameof(GetLabel)} - {ribbonControl?.Id}: '{label}'. Elapsed {sw.Elapsed}");
				return label;
			}
			catch (Exception ex)
			{
				logger.Error($"{nameof(GetLabel)} - {ribbonControl?.Id} error after {sw.Elapsed}:");
				logger.Error(ex);
				return ribbonControl?.Id;
			}
		}
		public Bitmap GetImage(string imageName)
		{
			var sw = Stopwatch.StartNew();
			logger.Debug($"{nameof(GetImage)} - {imageName}");
			try
			{
				var bitmap = (Bitmap)RibbonIcons.ResourceManager.GetObject(imageName);
				logger.Debug($"{nameof(GetImage)} - {imageName}: '{bitmap?.Size}'. Elapsed {sw.Elapsed}");
				return bitmap;
			}
			catch (Exception ex)
			{
				logger.Error($"{nameof(GetImage)} - {imageName} error after {sw.Elapsed}:");
				logger.Error(ex);
				return null;
			}
		}
		public string GetSupertip(IRibbonControl ribbonControl)
		{
			var sw = Stopwatch.StartNew();
			logger.Debug($"{nameof(GetSupertip)} - {ribbonControl?.Id}");
			try
			{
				var superTip = RibbonSupertips.ResourceManager.GetString(ribbonControl.Id);
				logger.Debug($"{nameof(GetSupertip)} - {ribbonControl?.Id}: '{superTip}'. Elapsed {sw.Elapsed}");
				return superTip;
			}
			catch (Exception ex)
			{
				logger.Error($"{nameof(GetSupertip)} - {ribbonControl?.Id} error after {sw.Elapsed}:");
				logger.Error(ex);
				return ribbonControl?.Id;
			}
		}

		private T GetService<T>()
		{
			return Container.Resolve<T>();
		}

		private static string GetResourceText()
		{
			var asm = typeof(NavfertyRibbon).Assembly;

			using (var stream = asm.GetManifestResourceStream("NavfertyExcelAddIn.NavfertyRibbon.xml"))
			using (var resourceReader = new StreamReader(stream))
			{
				return resourceReader.ReadToEnd();
			}
		}

		private Range GetSelectionOrUsedRange(Worksheet activeSheet)
		{
			var selection = (Range)App.Selection;

			if (activeSheet == null || selection == null)
				return null;

			var allCells = activeSheet.Cells.Rows.EntireRow;
			var isEntireSheetSelected = allCells.GetRelativeAddress() == selection.GetRelativeAddress();
			return isEntireSheetSelected
				? activeSheet.UsedRange
				: selection;
		}

		[SuppressMessage("Design", "CA1063:Implement IDisposable Correctly", Justification = "<Pending>")]
		public void Dispose()
		{
			form?.Dispose();
			Container.Dispose();
		}
		#endregion
	}
}
