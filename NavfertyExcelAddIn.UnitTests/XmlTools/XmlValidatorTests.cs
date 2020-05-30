using System;
using System.Reflection;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.XmlTools;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.XmlTools
{
	[TestClass]
	public class XmlValidatorTests : TestsBase
	{
		private Mock<IDialogService> dialogService;
		private Mock<IXsdSchemaValidator> xsdSchemaValidator;
		private Mock<Application> app;
		private Mock<Range> range;
		private XmlValidator validator;

		[TestInitialize]
		public void BeforeEachTest()
		{
			dialogService = new Mock<IDialogService>(MockBehavior.Strict);
			xsdSchemaValidator = new Mock<IXsdSchemaValidator>(MockBehavior.Strict);

			SetupApplicationForAddWorkbook();

			validator = new XmlValidator(dialogService.Object, xsdSchemaValidator.Object);
		}

		[TestMethod]
		public void Validate_XmlNotSelected()
		{
			dialogService
				.Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
				.Returns(Array.Empty<string>());


			validator.Validate(app.Object);

			dialogService.VerifyAll();
		}

		[TestMethod]
		public void Validate_XsdNotSelected()
		{
			dialogService
				.Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
				.Returns(new[] { "file.xml" });

			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
				.Returns(Array.Empty<string>());


			validator.Validate(app.Object);

			dialogService.VerifyAll();
		}

		[TestMethod]
		public void Validate_FilesNotExist_Throws()
		{
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
				.Returns(new[] { "file.xsd" });
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
				.Returns(new[] { "file.xml" });


			var ex = Assert.ThrowsException<ArgumentException>(() => validator.Validate(app.Object));

			Assert.AreEqual("One or more files not found", ex.Message);
			dialogService.VerifyAll();
		}


		[TestMethod]
		public void Validate_NoErrors_SuccessMessageShown()
		{
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
				.Returns(new[] { GetFilePath("XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd") });
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
				.Returns(new[] { GetFilePath("XmlTools/Samples/Прибыль_Correct.xml") });

			xsdSchemaValidator
				.Setup(x => x.Validate(It.IsAny<string>(), It.IsAny<string[]>()))
				.Returns(Array.Empty<XmlValidationError>());

			dialogService
				.SetupSequence(x => x.ShowInfo(It.IsAny<string>()));

			SetCulture();

			validator.Validate(app.Object);

			dialogService.VerifyAll();
			dialogService.Verify(x => x.ShowInfo(It.Is<string>(s => s.Contains("Successfully"))));
		}

		[TestMethod]
		public void Validate_HasErrors_MessageShown()
		{
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
				.Returns(new[] { GetFilePath("XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd") });
			dialogService
				.SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
				.Returns(new[] { GetFilePath("XmlTools/Samples/Прибыль_Correct.xml") });

			var error = new XmlValidationError(ValidationErrorSeverity.Error, "message", "value", "element name");
			xsdSchemaValidator
				.Setup(x => x.Validate(It.IsAny<string>(), It.IsAny<string[]>()))
				.Returns(new[] { error });

			SetCulture();

			validator.Validate(app.Object);

			range.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(s => s.Contains("element name"))));
			range.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Exactly(8));
		}

		private void SetupApplicationForAddWorkbook()
		{
			app = new Mock<Application>(MockBehavior.Strict);
			var wbs = new Mock<Workbooks>(MockBehavior.Strict);
			var wb = new Mock<Workbook>(MockBehavior.Strict);
			var wss = new Mock<Sheets>(MockBehavior.Strict);
			var ws = new Mock<Worksheet>(MockBehavior.Strict);
			//wss.Setup(x => x[1]).re
			var rangeBuilder = new RangeBuilder().WithSetValue();
			range = rangeBuilder.MockObject;
			range.Setup(x => x.get_Value(It.IsAny<object>()));
			range.Setup(x => x.set_Value(It.IsAny<object>(), It.IsAny<object>()));

			ws.Setup(x => x.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(rangeBuilder.Build());
			wb.Setup(x => x.Worksheets).Returns(wss.Object);
			wss.Setup(x => x[1]).Returns(ws.Object);
			wbs.Setup(x => x.Add(Missing.Value)).Returns(wb.Object);
			app.Setup(x => x.Workbooks).Returns(wbs.Object);
		}
	}
}
