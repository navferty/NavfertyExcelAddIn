using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

public class CaseTogglerTests : TestsBase
{
	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetRangeExtentionsStub();
	}

	[Test]
	public void ToggleCase_SingleCell_Converted()
	{
		var rangeBuilder = new RangeBuilder()
			.WithEnumrableRanges(new[] { new RangeBuilder().WithSingleValue("abc").Build() })
			.WithWorksheet()
			.WithAreas()
			.WithSingleValue("abc")
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var caseToggler = new CaseToggler();

		caseToggler.ToggleCase(selection);


		var expected = "ABC";
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => expected == v)));
	}

	[Test]
	public void ToggleCase_AllItemsConverted()
	{
		var values = new object?[,]
		{
			{ "abc", "def", "ghi" },
			{ "jkl", "123", 123d },
			{ "", null, (int)CVErrEnum.ErrNA }
		};
		var firstCell = new RangeBuilder().WithSingleValue(values[0, 0]).Build();

		var rangeBuilder = new RangeBuilder()
			.WithEnumrableRanges(new[] { firstCell })
			.WithMultipleValue(values)
			.WithWorksheet()
			.WithAreas()
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var caseToggler = new CaseToggler();

		caseToggler.ToggleCase(selection);


		var expected = new object?[,]
		{
			{ "ABC", "DEF", "GHI" },
			{ "JKL", "123", 123d },
			{ "", null, (int)CVErrEnum.ErrNA }
		};
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object?[,]>(v => AssertAssignedValue(expected, v).GetAwaiter().GetResult())));
	}

	[Test]
	[Arguments("abc", "ABC")]
	[Arguments("ABC", "Abc")]
	[Arguments("Abc", "abc")]

	[Arguments("A", "a")]
	[Arguments("a", "A")]
	public void ToggleCase_ConvertedRightOrder(string initial, string result)
	{
		var firstCell = new RangeBuilder().WithSingleValue(initial).Build();
		var rangeBuilder = new RangeBuilder()
				.WithEnumrableRanges(new[] { firstCell })
				.WithSingleValue(initial)
				.WithWorksheet()
				.WithAreas()
				.WithSetValue();
		var selection = rangeBuilder.Build();
		var caseToggler = new CaseToggler();

		caseToggler.ToggleCase(selection);


		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => result == v)));
	}

	[Test]
	public void ToggleCase_NoCellsWithLetters_NoChanges()
	{
		var values = new object?[,]
		{
			{ "", "123", 123d },
			{ "", null, (int)CVErrEnum.ErrNA }
		};
		var cells = new[]
		{
			new RangeBuilder().WithSingleValue("").Build(),
			new RangeBuilder().WithSingleValue("123").Build(),
			new RangeBuilder().WithSingleValue(123d).Build(),
			new RangeBuilder().WithSingleValue("").Build(),
			new RangeBuilder().WithSingleValue(null).Build(),
			new RangeBuilder().WithSingleValue((int)CVErrEnum.ErrNA).Build()
		};
		var rangeBuilder = new RangeBuilder()
			.WithEnumrableRanges(cells)
			.WithMultipleValue(values)
			.WithWorksheet()
			.WithAreas()
			.WithSetValue();

		var selection = rangeBuilder.Build();
		var caseToggler = new CaseToggler();

		caseToggler.ToggleCase(selection);

		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Never);
	}
	private static async Task<bool> AssertAssignedValue(object?[,] expected, object?[,] value)
	{
		await Assert.That(value).IsEquivalentTo(expected);
		return true;
	}
}
