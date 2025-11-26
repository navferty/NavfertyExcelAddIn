using Autofac;
using Autofac.Extras.DynamicProxy;

using Navferty.Common;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.DataValidation;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.SqliteExport;
using NavfertyExcelAddIn.StringifyNumerics;
using NavfertyExcelAddIn.Transliterate;
using NavfertyExcelAddIn.Undo;
using NavfertyExcelAddIn.UnprotectWorkbook;
using NavfertyExcelAddIn.WorksheetCellsEditing;
using NavfertyExcelAddIn.WorksheetProtectUnprotect;
using NavfertyExcelAddIn.XmlTools;

namespace NavfertyExcelAddIn
{
	public static class Registry
	{
		public static IContainer CreateContainer()
		{
			var builder = new ContainerBuilder();

			builder.RegisterType<DialogService>()
				.As<IDialogService>();

			builder.RegisterType<CellsValueValidator>()
				.As<ICellsValueValidator>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<ValidatorFactory>()
				.As<IValidatorFactory>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<CellsToMarkdownReader>()
				.As<ICellsToMarkdownReader>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<DuplicatesHighlighter>()
				.As<IDuplicatesHighlighter>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<EmptySpaceTrimmer>()
				.As<IEmptySpaceTrimmer>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<ConditionalFormatFixer>()
				.As<IConditionalFormatFixer>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<CaseToggler>()
				.As<ICaseToggler>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<CellsUnmerger>()
				.As<ICellsUnmerger>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<WbUnprotector>()
				.As<IWbUnprotector>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<WsProtectorUnprotector>()
				.As<IWsProtectorUnprotector>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<ErrorFinder>()
				.As<IErrorFinder>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<ParseNumerics.NumericParserService>()
				.As<INumericParser>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<RussianNumericStringifier>()
				.Keyed<INumericStringifier>(SupportedCulture.Russian)
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<EnglishNumericStringifier>()
				.Keyed<INumericStringifier>(SupportedCulture.English)
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<FrenchNumericStringifier>()
				.Keyed<INumericStringifier>(SupportedCulture.French)
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<CyrillicLettersReplacer>()
				.As<ICyrillicLettersReplacer>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<Transliterator>()
				.As<ITransliterator>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<XmlValidator>()
				.As<IXmlValidator>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<XsdSchemaValidator>()
				.As<IXsdSchemaValidator>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<XmlSampleCreator>()
				.As<IXmlSampleCreator>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<SqliteExportOptionsFormProvider>()
				.As<ISqliteExportOptionsProvider>();

			builder.RegisterType<SqliteExporter>()
				.As<ISqliteExporter>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<Web.WebToolsBuilder>()
				.As<Web.IWebTools>()
				.EnableInterfaceInterceptors()
				.InterceptedBy(typeof(ExceptionLogger));

			builder.RegisterType<ExceptionLogger>();

			builder.RegisterType<UndoManager>()
				.SingleInstance();

			return builder.Build();
		}
	}
}
