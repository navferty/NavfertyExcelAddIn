using System;
using System.IO;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

using NLog;

namespace NavfertyExcelAddIn.XmlTools
{
	public class XmlValidator : IXmlValidator
	{
		private readonly IDialogService dialogService;
		private readonly IXsdSchemaValidator xsdValidator;

		private readonly Logger logger = LogManager.GetCurrentClassLogger();

		public XmlValidator(IDialogService dialogService, IXsdSchemaValidator xsdValidator)
		{
			this.dialogService = dialogService;
			this.xsdValidator = xsdValidator;
		}

		public void Validate(Application excelApp)
		{
			var xmlFileName = dialogService.AskForFiles(false, FileType.Xml).FirstOrDefault();

			if (string.IsNullOrEmpty(xmlFileName))
			{
				return;
			}

			// allow multiple xsd files as one schema
			var xsdFileNames = dialogService.AskForFiles(true, FileType.Xsd);

			if (string.IsNullOrEmpty(xsdFileNames.FirstOrDefault()))
			{
				return;
			}

			if (!File.Exists(xmlFileName) || xsdFileNames.Any(path => !File.Exists(path)))
			{
				throw new ArgumentException("One or more files not found");
			}

			logger.Debug($"Try to validate {xmlFileName} by {string.Join(", ", xsdFileNames)}");

			var errors = xsdValidator.Validate(xmlFileName, xsdFileNames);

			if (!errors.Any())
			{
				dialogService.ShowInfo(UIStrings.SuccessfullyValidatedMessage);
				return;
			}

			logger.Debug($"{errors.Count} errors");

			var ws = (Worksheet)excelApp.Workbooks.Add().Worksheets[1];
			((Range)ws.Cells[1, 1]).Value = UIStrings.XmlValidationReport_Severity;
			((Range)ws.Cells[1, 2]).Value = UIStrings.XmlValidationReport_ElementName;
			((Range)ws.Cells[1, 3]).Value = UIStrings.XmlValidationReport_Value;
			((Range)ws.Cells[1, 4]).Value = UIStrings.XmlValidationReport_Message;

			var i = 2;
			foreach (var error in errors)
			{
				((Range)ws.Cells[i, 1]).Value = error.Severity.ToString();
				((Range)ws.Cells[i, 2]).Value = error.ElementName;
				((Range)ws.Cells[i, 3]).Value = error.Value;
				((Range)ws.Cells[i, 4]).Value = error.Message;
				i++;
			}
		}
	}
}
