using Navferty.Common;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnitTests.Builders;

namespace NavfertyExcelAddIn.UnitTests.FindFormulaErrors;

public class ErrorFinderTests : TestsBase
{
	[Test]
	public async Task Range_NoErrors()
	{
		var errorFinder = new ErrorFinder();

		var values = new object?[,] { { 1d, "1", "abc" }, { "123.123", "321 , 321", null } };
		var range = new RangeBuilder().WithWorksheet().WithSingleValue(values).Build();

		var result = errorFinder.GetAllErrorCells(range).ToArray();

		await Assert.That(result).IsEmpty();
	}

	[Test]
	public async Task Range_WithErrors()
	{
		var errorFinder = new ErrorFinder();

		var values = new object?[,] { { 1d, "1", "abc" }, { (int)CVErrEnum.ErrGettingData, "321 , 321", null } };
		var range = new RangeBuilder().WithWorksheet().WithCells().WithSingleValue(values).Build();
		SetRangeExtentionsStub();

		var result = errorFinder.GetAllErrorCells(range).ToArray();

		await Assert.That(result).Count().EqualTo(1);
		await Assert.That(result.First().ErrorMessage).IsEqualTo(CVErrEnum.ErrGettingData.GetEnumDescription());
	}

	[Test]
	public async Task SingleValue_WithoutError()
	{
		var errorFinder = new ErrorFinder();

		var range = new RangeBuilder().WithWorksheet().WithSingleValue("1").Build();

		var result = errorFinder.GetAllErrorCells(range).ToArray();

		await Assert.That(result).IsEmpty();
	}

	[Test]
	public async Task SingleValue_WithError()
	{
		var errorFinder = new ErrorFinder();

		var range = new RangeBuilder().WithWorksheet().WithSingleValue((int)CVErrEnum.ErrNA).Build();

		SetRangeExtentionsStub();


		var result = errorFinder.GetAllErrorCells(range).ToArray();

		await Assert.That(result).Count().EqualTo(1);
		await Assert.That(result.First().ErrorMessage).IsEqualTo(CVErrEnum.ErrNA.GetEnumDescription());
	}

	[Test]
	public async Task NoValues()
	{
		var errorFinder = new ErrorFinder();

		var range = new RangeBuilder().WithWorksheet().WithSingleValue(null).Build();

		var result = errorFinder.GetAllErrorCells(range).ToArray();

		await Assert.That(result).IsEmpty();
	}
}
