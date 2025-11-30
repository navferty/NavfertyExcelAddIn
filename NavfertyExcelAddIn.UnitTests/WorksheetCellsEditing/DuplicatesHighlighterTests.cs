using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

// Disable parallelization due to RangeExtensionsStub usage in static context
[NotInParallel(RangeExtensionsStub.ParallelizationContstraintKey)]
public class DuplicatesHighlighterTests : TestsBase
{
	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetRangeExtentionsStub();
	}

	[Test]
	public async Task HighlightDuplicates_ThreeItems_Success()
	{
		var cells = new[]
		{
			new RangeBuilder().WithSingleValue("abc").Build(),
			new RangeBuilder().WithSingleValue("a").Build(),
			new RangeBuilder().WithSingleValue("b").Build(),
			new RangeBuilder().WithSingleValue("c").Build(),
			new RangeBuilder().WithSingleValue("a").Build(),
			new RangeBuilder().WithSingleValue("b").Build(),
			new RangeBuilder().WithSingleValue("qwe").Build(),
			new RangeBuilder().WithSingleValue(2d).Build(),
			new RangeBuilder().WithSingleValue(3d).Build(),
			new RangeBuilder().WithSingleValue(2d).Build()
		};
		var selection = new RangeBuilder().WithEnumrableRanges(cells).Build();
		var duplicatesHighlighter = new DuplicatesHighlighter();

		duplicatesHighlighter.HighlightDuplicates(selection);

		await Assert.That(RangeExtensionsStub.SetColorInvocations.Count()).IsEqualTo(3);
	}

	[Test]
	public async Task HighlightDuplicates_ManyItemsWithSameColors_Success()
	{
		var cells = Enumerable.Range(1, 1000).Select(x => new RangeBuilder().WithSingleValue((x % 100).ToString()).Build());
		var selection = new RangeBuilder().WithEnumrableRanges(cells).Build();
		var duplicatesHighlighter = new DuplicatesHighlighter();

		duplicatesHighlighter.HighlightDuplicates(selection);

		// all 57 colors
		await Assert.That(RangeExtensionsStub.SetColorInvocations.Count()).IsEqualTo(57);
	}

	[Test]
	public async Task HighlightDuplicates_SomeCellsAreNull_Success()
	{
		var cells = new[]
		{
			new RangeBuilder().WithSingleValue(1).Build(),
			new RangeBuilder().WithSingleValue(2).Build(),
			new RangeBuilder().WithSingleValue(3).Build(),
			null,
			new RangeBuilder().WithSingleValue(3).Build(),
			new RangeBuilder().WithSingleValue(4).Build(),
		};
		var selection = new RangeBuilder().WithEnumrableRanges(cells).Build();
		var duplicatesHighlighter = new DuplicatesHighlighter();

		duplicatesHighlighter.HighlightDuplicates(selection);

		await Assert.That(RangeExtensionsStub.SetColorInvocations.Count()).IsEqualTo(1);
	}
}
