using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.DataValidation;
using NavfertyExcelAddIn.UnitTests.Builders;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.DataValidation
{
	[TestClass]
	public class CellsValueValidatorTests : TestsBase
	{
		private Mock<IValidator> validator;

		private CellsValueValidator cellsValidator;

		[TestInitialize]
		public void BeforeEachTest()
		{
			SetRangeExtentionsStub();

			validator = new Mock<IValidator>(MockBehavior.Strict);

			var validatorFactory = new Mock<IValidatorFactory>(MockBehavior.Strict);
			validatorFactory.Setup(x => x.CreateValidator(ValidationType.Xml)).Returns(validator.Object);

			cellsValidator = new CellsValueValidator(validatorFactory.Object);
		}

		[TestMethod]
		public void Validate_All_Success()
		{
			validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Success);
			var selection = new RangeBuilder().WithEnumrableRanges().WithWorksheet().Build();

			var result = cellsValidator.Validate(selection, ValidationType.Xml);

			Assert.AreEqual(0, result.Count);
		}

		[TestMethod]
		public void Validate_All_Fail()
		{
			validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Fail(string.Empty));
			var selection = new RangeBuilder().WithEnumrableRanges().WithWorksheet().Build();

			var result = cellsValidator.Validate(selection, ValidationType.Xml);

			Assert.AreEqual(10, result.Count);
		}

		[TestMethod]
		public void Validate_Success_And_Fail()
		{
			var cells = new[]
			{
				new RangeBuilder().WithSingleValue(true).Build(),
				new RangeBuilder().WithSingleValue(true).Build(),
				new RangeBuilder().WithSingleValue(false).Build(),
				new RangeBuilder().WithSingleValue("").Build()
			};
			var selection = new RangeBuilder().WithEnumrableRanges(cells).WithWorksheet().Build();

			validator
				.Setup(x => x.CheckValue(It.IsAny<object>()))
				.Returns((bool x) => x ? ValidationResult.Success : ValidationResult.Fail("Fail =("));

			var result = cellsValidator.Validate(selection, ValidationType.Xml);

			Assert.AreEqual(1, result.Count);
			validator.Verify(x => x.CheckValue(It.IsAny<object>()), Times.Exactly(3));
		}

		[TestMethod]
		public void Validate_NoCells()
		{
			var selection = new RangeBuilder().WithEnumrableRanges(Array.Empty<Range>()).Build();

			var result = cellsValidator.Validate(selection, ValidationType.Xml);

			Assert.AreEqual(0, result.Count);
		}
	}
}
