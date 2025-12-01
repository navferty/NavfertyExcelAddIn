using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

[NotInParallel(RangeExtensionsStub.ParallelizationContstraintKey)]
public class EmptySpaceTrimmerTests : TestsBase
{
	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetRangeExtentionsStub();
	}

	[Test]
	public void TrimSpaces_SingleCell_Converted()
	{
		var value = "a   \r\nb     c\t";
		var rangeBuilder = new RangeBuilder()
			.WithWorksheet()
			.WithAreas()
			.WithSingleValue(value)
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var emptySpaceTrimmer = new EmptySpaceTrimmer();

		emptySpaceTrimmer.TrimExtraSpaces(selection);

		var expected = "a b c";
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => expected == v)));
	}

	[Test]
	public void TrimSpaces_AllItemsTrimmed()
	{
		var values = new object?[,]
		{
			{ "abc ", " def", "  ghi\t" },
			{ "jk\r\nl", "   \r\n   \t", "   " },
			{ "", null, (int)CVErrEnum.ErrNA }
		};
		var rangeBuilder = new RangeBuilder()
			.WithWorksheet()
			.WithAreas()
			.WithMultipleValue(values)
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var emptySpaceTrimmer = new EmptySpaceTrimmer();

		emptySpaceTrimmer.TrimExtraSpaces(selection);

		var expected = new object?[,]
		{
			{ "abc", "def", "ghi" },
			{ "jk l", null, null },
			{ null, null, (int)CVErrEnum.ErrNA }
		};
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v).GetAwaiter().GetResult())));
	}

	[Test]
	public void RemoveAllSpaces_EmptyValue_Null()
	{
		var value = "   \r\n     \t";
		var rangeBuilder = new RangeBuilder()
			.WithWorksheet()
			.WithAreas()
			.WithSingleValue(value)
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var emptySpaceTrimmer = new EmptySpaceTrimmer();

		emptySpaceTrimmer.RemoveAllSpaces(selection);

		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => v == null)));
	}

	[Test]
	public void RemoveAllSpaces_AllValues_NoSpacesLeft()
	{
		var values = new object?[,]
		{
			{ "abc ", " def", "  g h i\t" },
			{ "jk\r\nl", "   \r\n   \t", "   " },
			{ "", null, (int)CVErrEnum.ErrNA }
		};
		var rangeBuilder = new RangeBuilder()
			.WithWorksheet()
			.WithAreas()
			.WithMultipleValue(values)
			.WithSetValue();
		var selection = rangeBuilder.Build();
		var emptySpaceTrimmer = new EmptySpaceTrimmer();

		emptySpaceTrimmer.RemoveAllSpaces(selection);

		var expected = new object?[,]
		{
			{ "abc", "def", "ghi" },
			{ "jkl", null, null },
			{ null, null, (int)CVErrEnum.ErrNA }
		};
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
		rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v).GetAwaiter().GetResult())));
	}

	private async Task<bool> AssertAssignedValue(object?[,] expected, object?[,] value)
	{
		await Assert.That(value).IsEquivalentTo(expected);
		return true;
	}
}
