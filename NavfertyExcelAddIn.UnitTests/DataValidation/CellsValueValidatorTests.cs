using Moq;

using NavfertyExcelAddIn.DataValidation;
using NavfertyExcelAddIn.UnitTests.Builders;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.DataValidation;

[NotInParallel(RangeExtensionsStub.ParallelizationContstraintKey)]
public class CellsValueValidatorTests : TestsBase
{
	private readonly Mock<IValidator> validator = new(MockBehavior.Strict);
	private readonly Mock<IValidatorFactory> validatorFactory = new(MockBehavior.Strict);

	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetRangeExtentionsStub();
		validatorFactory.Setup(x => x.CreateValidator(ValidationType.Xml)).Returns(validator.Object);
	}

	[Test]
	public async Task Validate_All_Success()
	{
		validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Success);
		var selection = new RangeBuilder().WithEnumrableRanges().WithWorksheet().Build();
		var cellsValidator = new CellsValueValidator(validatorFactory.Object);

		var result = cellsValidator.Validate(selection, ValidationType.Xml);

		await Assert.That(result).IsEmpty();
	}

	[Test]
	public async Task Validate_All_Fail()
	{
		validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Fail(string.Empty));
		var selection = new RangeBuilder().WithEnumrableRanges().WithWorksheet().Build();
		var cellsValidator = new CellsValueValidator(validatorFactory.Object);

		var result = cellsValidator.Validate(selection, ValidationType.Xml);

		await Assert.That(result).Count().EqualTo(10);
	}

	[Test]
	public async Task Validate_Success_And_Fail()
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
		var cellsValidator = new CellsValueValidator(validatorFactory.Object);

		var result = cellsValidator.Validate(selection, ValidationType.Xml);

		await Assert.That(result).Count().EqualTo(1);
		validator.Verify(x => x.CheckValue(It.IsAny<object>()), Times.Exactly(3));
	}

	[Test]
	public async Task Validate_NoCells()
	{
		var selection = new RangeBuilder().WithEnumrableRanges(Array.Empty<Range>()).Build();
		var cellsValidator = new CellsValueValidator(validatorFactory.Object);

		var result = cellsValidator.Validate(selection, ValidationType.Xml);

		await Assert.That(result).IsEmpty();
	}
}
